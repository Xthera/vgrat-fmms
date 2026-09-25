#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings FAILED-FUND RECOVERY 2

PURPOSE
=======

Recover ONLY funds that still failed during Recovery 1:

    scripts/test_prudential_holdings.py

Recovery 1 remains untouched.

RECOVERY SOURCE
===============

1. output_holdings_recovery/run_summary.json
       |
       v
   stillFailedFundsDetail
       |
       v
2. Funds Links.xlsm
       |
       v
   authoritative Excel URL + PruAccess name
       |
       v
3. Official Prudential Singapore fund page
       |
       v
4. Official Prudential factsheet PDF
       |
       v
5. Normal PDF extraction
       |
       v
6. Primary holdings parser
       |
       v
7. Fallback holdings parser
       |
       v
8. Spatial PDF recovery
       |
       v
9. Exact official-PDF re-download
       |
       v
10. Exact extraction/signature verification


IMPORTANT
=========

This script is a SEPARATE recovery script.

It does NOT modify:

    scripts/test_prudential_holdings.py

It does NOT expand the universe.

It does NOT use PruAccess to discover holdings.

It does NOT use third-party holdings sources.

It does NOT fabricate holdings.

It does NOT estimate percentages.

It does NOT use fuzzy matching.

It does NOT use proximity matching between arbitrary names and
percentages.

It does NOT silently accept a different result during final verification.


HARD RULES
==========

- Recovery 1 stillFailedFundsDetail controls the Recovery 2 universe.
- Funds Links.xlsm controls the authoritative URL and Excel row.
- Only official Prudential Singapore URLs are accepted.
- Only official Prudential Singapore factsheets are accepted.
- No third-party holdings sources.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No forced ten holdings.
- Fewer than ten published holdings is valid.
- Duplicate holding names are allowed.
- Duplicate holding percentages are allowed.
- Published order must be preserved.
- Fixed-income coupon percentages may occur inside holding names.
- The last percentage on a logical fixed-income line is the portfolio
  weight when fallback parsing is required.
- A hyphenated PDF line break is rejoined before percentage parsing.
- Spatial recovery MUST still pass the reconstructed section through
  the fallback parser.
- Spatial recovery does NOT independently pair names and percentages
  using coordinate proximity.
- Final verification downloads the official PDF again.
- Final verification runs the extraction engine again.
- Final verification requires the exact same:
      rank
      holding name
      weight
  signature.
- If exact verification fails, the fund remains FAILED.
- A recovery result never overwrites the baseline result.
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

RECOVERY_1_RUN_SUMMARY_FILE = Path(
    "output_holdings_recovery/run_summary.json"
)

RECOVERY_OUTPUT_DIR = Path(
    "output_holdings_recovery_2"
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
# SPATIAL RECOVERY
# =============================================================================

SPATIAL_MIN_NAME_LENGTH = 3

SPATIAL_MAX_LINES_AFTER_HEADING = 80


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
# BASELINE
# =============================================================================

def load_recovery_1_run_summary() -> dict:

    if not RECOVERY_1_RUN_SUMMARY_FILE.exists():

        raise FileNotFoundError(
            "Recovery 1 source run summary not found: "
            f"{RECOVERY_1_RUN_SUMMARY_FILE}"
        )

    data = json.loads(
        RECOVERY_1_RUN_SUMMARY_FILE.read_text(
            encoding="utf-8"
        )
    )

    if not isinstance(
        data,
        dict,
    ):

        raise RuntimeError(
            "Baseline run summary is not a JSON object."
        )

    failed = data.get(
        "stillFailedFundsDetail"
    )

    if not isinstance(
        failed,
        list,
    ):

        raise RuntimeError(
            "Baseline run summary does not contain "
            "a valid stillFailedFundsDetail list."
        )

    return data


# =============================================================================
# EXCEL MASTER UNIVERSE
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
            "No populated Prudential URLs were found "
            "in Excel Column A."
        )

    return funds


def build_recovery_universe(
    recovery_1: dict,
    excel_funds: dict[int, dict],
) -> list[dict]:

    failed_entries = recovery_1.get(
        "stillFailedFundsDetail"
    )

    if not isinstance(
        failed_entries,
        list,
    ):

        raise RuntimeError(
            "Recovery 1 stillFailedFundsDetail is invalid."
        )

    recovery = []

    seen_rows = set()

    for failure in failed_entries:

        if not isinstance(
            failure,
            dict,
        ):

            raise RuntimeError(
                "A stillFailedFundsDetail entry is not an object."
            )

        if "excelRow" not in failure:

            raise RuntimeError(
                "A Recovery 1 failure entry has no excelRow."
            )

        try:

            excel_row = int(
                failure[
                    "excelRow"
                ]
            )

        except Exception as error:

            raise RuntimeError(
                "Recovery 1 failure contains an invalid excelRow: "
                f"{failure.get('excelRow')}"
            ) from error

        if excel_row in seen_rows:

            raise RuntimeError(
                f"Duplicate Recovery 1 failed excelRow: {excel_row}"
            )

        seen_rows.add(
            excel_row
        )

        if excel_row not in excel_funds:

            raise RuntimeError(
                "Recovery 1 failed row "
                f"{excel_row} does not exist in Excel."
            )

        excel_fund = excel_funds[
            excel_row
        ]

        baseline_url = clean_text(
            failure.get(
                "prudentialUrl"
            )
        )

        baseline_name = clean_text(
            failure.get(
                "pruAccessName"
            )
            or
            failure.get(
                "excelPruAccessName"
            )
        )

        excel_url = clean_text(
            excel_fund[
                "prudentialUrl"
            ]
        )

        excel_name = clean_text(
            excel_fund.get(
                "pruAccessName"
            )
        )

        if (
            baseline_url
            and
            baseline_url
            !=
            excel_url
        ):

            raise RuntimeError(
                "Recovery 1 URL does not match Excel for "
                f"row {excel_row}."
            )

        if (
            baseline_name
            and
            excel_name
            and
            baseline_name
            !=
            excel_name
        ):

            raise RuntimeError(
                "Recovery 1 PruAccess name does not match Excel for "
                f"row {excel_row}."
            )

        recovery.append(
            {
                "excelRow":
                    excel_row,

                "prudentialUrl":
                    excel_url,

                "pruAccessName":
                    excel_name,

                "baselineFailure":
                    failure,
            }
        )

    return recovery


# =============================================================================
# FACTSHEET DISCOVERY
# =============================================================================

def find_factsheet_url(
    page,
) -> str:

    anchors = page.locator(
        "a"
    )

    count = anchors.count()

    candidates = []

    for index in range(
        count
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
            "Could not find an official Prudential factsheet link."
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
            "Factsheet HTTP status was "
            f"{response.status}."
        )

    body = response.body()

    if not body:

        raise RuntimeError(
            "Factsheet response contained zero bytes."
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
            f"Content-Type={content_type}"
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
            "Factsheet PDF has zero pages."
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
                "Failed to extract PDF text from page "
                f"{page_number}: {error}"
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
            "Factsheet PDF contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


# =============================================================================
# PDF LINES
# =============================================================================

def pdf_lines(
    text: str,
) -> list[str]:

    result = []

    for raw_line in text.splitlines():

        line = clean_text(
            raw_line
        )

        if line:
            result.append(
                line
            )

    return result


# =============================================================================
# FACTSHEET DATES
# =============================================================================

def extract_data_as_at(
    text: str,
) -> str | None:

    patterns = [

        re.compile(
            r"\bdata\s+as\s+at\s*"
            r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})\b",
            re.IGNORECASE,
        ),

        re.compile(
            r"\ball\s+data\s+as\s+at\s*"
            r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})\b",
            re.IGNORECASE,
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


def extract_document_date(
    text: str,
) -> str | None:

    pattern = re.compile(
        r"\b("
        r"January|February|March|April|May|June|July|"
        r"August|September|October|November|December"
        r")\s+([0-9]{4})\b",
        re.IGNORECASE,
    )

    for line in pdf_lines(
        text
    )[:50]:

        match = pattern.search(
            line
        )

        if match:

            return clean_text(
                match.group(0)
            )

    return None


# =============================================================================
# HOLDINGS SECTION
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
# NAME CLEANING
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

    name = name.strip(
        " -|"
    )

    if normalize_text(
        name
    ) in {
        "",
        "none",
        "null",
        "-",
        "—",
    }:

        return ""

    return name


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

    combined = " ".join(
        cleaned
    )

    combined = re.sub(
        r"\s+",
        " ",
        combined,
    ).strip()

    return clean_holding_name(
        combined
    )


# =============================================================================
# PERCENTAGES
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


def find_percentage_in_line(
    line: str,
) -> tuple[float, str, int, int] | None:

    matches = percentage_matches(
        line
    )

    if not matches:
        return None

    match = matches[0]

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
# RANKS
# =============================================================================

def extract_leading_rank(
    line: str,
) -> tuple[int | None, str]:

    text = clean_text(
        line
    )

    if not text:
        return (
            None,
            "",
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


# =============================================================================
# NOISE
# =============================================================================

def is_holding_header_or_noise(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return True

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
        "portfolio holdings",
    }:

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


# =============================================================================
# HYPHENATED LINE RECONSTRUCTION
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


# =============================================================================
# LOGICAL LINES
# =============================================================================

def build_logical_holding_lines(
    section_text: str,
) -> list[str]:

    raw_lines = pdf_lines(
        section_text
    )

    if not raw_lines:
        return []

    return merge_hyphenated_line_breaks(
        raw_lines
    )


# =============================================================================
# VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
    parser_name: str,
) -> list[dict]:

    if not holdings:

        raise RuntimeError(
            f"{parser_name} parser produced no holdings."
        )

    if len(holdings) > MAX_HOLDINGS:

        raise RuntimeError(
            f"{parser_name} parser produced more than "
            f"{MAX_HOLDINGS} holdings."
        )

    expected = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    ranks = [
        item[
            "rank"
        ]
        for item in holdings
    ]

    if ranks != expected:

        raise RuntimeError(
            f"{parser_name} parser ranks invalid. "
            f"Parsed={ranks}; Expected={expected}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding.get(
                "name"
            )
        )

        if not name:

            raise RuntimeError(
                f"{parser_name} parser produced an empty holding name."
            )

        weight = holding.get(
            "weightPercent"
        )

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                f"{parser_name} parser produced an invalid weight."
            )

        if not 0 <= weight <= 100:

            raise RuntimeError(
                f"{parser_name} parser produced weight outside 0-100."
            )

        holding[
            "name"
        ] = name

    return holdings


# =============================================================================
# PRIMARY PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:

    lines = build_logical_holding_lines(
        section_text
    )

    if not lines:

        raise RuntimeError(
            "Holdings section is empty."
        )

    holdings = []

    pending_fragments = []

    pending_rank = None

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

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            if pending_fragments:

                raise RuntimeError(
                    "Primary parser encountered a new rank "
                    "before the previous holding received a weight."
                )

            pending_rank = detected_rank

            line = remainder

            if not line:
                continue

        matches = percentage_matches(
            line
        )

        if len(matches) > 1:

            raise RuntimeError(
                "Primary parser found multiple percentages on "
                "one logical line."
            )

        percentage_info = (
            find_percentage_in_line(
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

            fragment = clean_text(
                line[:start]
            )

            if fragment:
                pending_fragments.append(
                    fragment
                )

            name = combine_holding_name_fragments(
                pending_fragments
            )

            if not name:

                raise RuntimeError(
                    "Primary parser found a percentage "
                    "without a holding name."
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
                        weight,

                    "weightText":
                        weight_text,
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

    return validate_holdings(
        holdings,
        "Primary",
    )


# =============================================================================
# FALLBACK PARSER
# =============================================================================

def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:

    lines = build_logical_holding_lines(
        section_text
    )

    if not lines:

        raise RuntimeError(
            "Fallback parser received an empty section."
        )

    holdings = []

    pending_fragments = []

    pending_rank = None

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

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            if pending_fragments:

                raise RuntimeError(
                    "Fallback parser encountered a new rank "
                    "before the previous holding received a weight."
                )

            pending_rank = detected_rank

            line = remainder

            if not line:
                continue

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

            fragment = clean_text(
                line[:start]
            )

            if fragment:
                pending_fragments.append(
                    fragment
                )

            name = combine_holding_name_fragments(
                pending_fragments
            )

            if not name:

                raise RuntimeError(
                    "Fallback parser found a percentage "
                    "without a holding name."
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
                        weight,

                    "weightText":
                        weight_text,
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

    return validate_holdings(
        holdings,
        "Fallback",
    )


# =============================================================================
# HOLDING SIGNATURE
# =============================================================================

def holding_signature(
    holdings: list[dict],
) -> list[tuple]:

    signature = []

    for holding in holdings:

        signature.append(
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
        )

    return signature


# =============================================================================
# SPATIAL PDF EXTRACTION
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


def _spatial_row_text(line: dict) -> str:
    """Return the text from one physical left-column visual row."""

    fragments = list(line.get("fragments", []))
    fragments.sort(key=lambda item: item["x0"])

    return clean_text(
        " ".join(
            fragment["text"]
            for fragment in fragments
        )
    )


def _is_spatial_nonholding_row(text: str) -> bool:
    """
    Reject obvious page/chart material from a spatial candidate.

    This is deliberately conservative. A row is not rejected merely because
    it contains a percentage; fixed-income holding rows legitimately contain
    coupon percentages plus the final portfolio weight.
    """

    normalized = normalize_text(text)

    if not normalized:
        return True

    if normalized in {
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
    }:
        return True

    if normalized.startswith("source:"):
        return True

    if normalized.startswith("important information"):
        return True

    if normalized.startswith("top 10 holdings"):
        return True

    return False


def _build_spatial_holdings_section(
    page_lines: list[dict],
    heading: dict,
    page_width: float,
) -> str:
    """
    Build a visual-order Top Holdings section from one page.

    IMPORTANT:
    ----------
    This function only reconstructs the physical left-hand holdings column.
    It does not parse or infer holdings itself.
    """

    left_edge = max(
        0.0,
        heading["x0"] - 15.0,
    )

    right_edge = _spatial_column_boundary(
        page_lines,
        heading,
        page_width,
    )

    selected = []

    # PDF coordinates normally increase upward, so visible content below the
    # heading has a smaller y coordinate.
    for line in page_lines:
        if line["y"] >= heading["y"] - 2.0:
            continue

        if line["y"] < heading["y"] - 520.0:
            continue

        cropped_fragments = []

        for fragment in line.get("fragments", []):
            x0 = float(fragment.get("x0", 0.0))

            if x0 < left_edge:
                continue

            if x0 >= right_edge:
                continue

            cropped_fragments.append(
                fragment
            )

        if not cropped_fragments:
            continue

        cropped_fragments.sort(
            key=lambda item: item["x0"]
        )

        cropped_text = clean_text(
            " ".join(
                fragment["text"]
                for fragment in cropped_fragments
            )
        )

        if not cropped_text:
            continue

        normalized = normalize_text(
            cropped_text
        )

        # A second visible Top Holdings heading means this candidate has
        # reached another section on the same page.
        if _is_spatial_top_holdings_heading(
            cropped_text
        ):
            break

        if is_holdings_end(
            cropped_text
        ):
            break

        if normalized in {
            "dividend history",
            "distribution history",
            "date",
            "frequency",
        }:
            continue

        cropped_line = dict(line)
        cropped_line["text"] = cropped_text
        cropped_line["fragments"] = cropped_fragments
        cropped_line["x0"] = min(
            fragment["x0"]
            for fragment in cropped_fragments
        )
        cropped_line["x1"] = max(
            fragment["x0"]
            for fragment in cropped_fragments
        )

        selected.append(
            cropped_line
        )

    selected.sort(
        key=lambda item: -item["y"]
    )

    return "\n".join(
        _spatial_row_text(line)
        for line in selected
        if _spatial_row_text(line)
    )


def _parse_spatial_holdings_rows(
    section_lines: list[dict],
) -> list[dict]:
    """
    Parse the physical Top Holdings table after PDF text has been grouped by
    visual baseline.

    IMPORTANT FOR PRUDENTIAL FIXED-INCOME FACTSHEETS
    ------------------------------------------------
    Some Prudential bond tables do not expose one complete holding as one
    pypdf text line.  A wrapped security can be emitted physically as:

        SEATRIUM FINANCIAL SERVICES PTE LTD
        2.95% 28-APR-2031                         1.6%

    or, depending on PDF object order:

        SEATRIUM FINANCIAL SERVICES PTE LTD
        2.95% 28-APR-2031
                                                1.6%

    The first percentage is the security coupon.  The final percentage is the
    portfolio weight.  Therefore this parser first reconstructs a *visual
    holding row* from consecutive left-column lines, then takes the LAST
    percentage belonging to that reconstructed row as the weight.

    This is not arbitrary nearest-neighbour pairing.  A detached weight is
    accepted only when it immediately follows a name/continuation group in
    the same Top Holdings table and remains inside the strict wrapped-row
    vertical band.  Unrelated percentage rows are rejected.
    """

    # The PDF page uses a two-column table.  The first eight rows are single
    # visual lines, while the Seatrium row is visibly wrapped.  Keep the
    # vertical grouping deliberately tight so the next holding cannot be
    # swallowed into the previous one.
    WRAPPED_ROW_MAX_GAP = 24.0

    holdings: list[dict] = []
    pending_name_fragments: list[str] = []
    pending_last_y: float | None = None

    percentage_pattern = re.compile(
        r"(?<![\d.])([0-9]+(?:\.[0-9]+)?)\s*%"
    )

    maturity_pattern = re.compile(
        r"\b\d{1,2}-[A-Za-z]{3}-\d{4}\b"
    )

    def commit(
        name: str,
        percentage: float,
        percentage_text: str,
    ) -> None:
        name = clean_holding_name(name)

        if not name:
            raise RuntimeError(
                "Spatial parser found a published holding percentage "
                "but no holding name could be reconstructed."
            )

        if len(holdings) >= MAX_HOLDINGS:
            raise RuntimeError(
                "Spatial parser produced more than 10 holdings."
            )

        holdings.append(
            {
                "rank": len(holdings) + 1,
                "name": name,
                "weightPercent": percentage,
                "weightText": percentage_text,
            }
        )

    def reset_pending() -> None:
        nonlocal pending_name_fragments, pending_last_y
        pending_name_fragments = []
        pending_last_y = None

    def append_pending(fragment: str, y: float) -> None:
        nonlocal pending_last_y

        fragment = clean_holding_fragment(fragment)
        if not fragment:
            return

        if pending_last_y is not None and abs(y - pending_last_y) > WRAPPED_ROW_MAX_GAP:
            # A new visual area has started.  Never carry an old name into it.
            pending_name_fragments.clear()

        pending_name_fragments.append(fragment)
        pending_last_y = y

    def is_coupon_continuation(text: str, matches: list[re.Match]) -> bool:
        """
        Recognise a fixed-income security continuation such as:

            2.95% 28-APR-2031

        It contains a percentage, but that percentage is part of the security
        name because a maturity date follows it.  It therefore MUST NOT be
        consumed as the portfolio weight.
        """

        if not matches:
            return False

        # Only the FINAL percentage matters here.  If the final percentage
        # itself is followed by a maturity date, it is part of the security
        # description.  In a complete row such as
        # ``2.375% 1-JUL-2039 5.7%``, the final 5.7% is after the maturity
        # date and is therefore the portfolio weight.
        last_match = matches[-1]
        return bool(maturity_pattern.search(text[last_match.end():]))

    # The section is already in visual top-to-bottom order, but sort again so
    # the parser never depends on dictionary insertion order.
    ordered = sorted(
        section_lines,
        key=lambda item: (-float(item["y"]), float(item.get("x0", 0.0))),
    )

    for line in ordered:
        text = clean_text(
            line.get("text") or _spatial_row_text(line)
        )

        if not text:
            continue

        if _is_spatial_nonholding_row(text):
            continue

        y = float(line["y"])
        matches = list(percentage_pattern.finditer(text))

        # ------------------------------------------------------------------
        # No percentage: wrapped security-name text.
        # ------------------------------------------------------------------
        if not matches:
            append_pending(text, y)
            continue

        # ------------------------------------------------------------------
        # A coupon + maturity continuation is still NAME text.
        # Example:
        #     2.95% 28-APR-2031
        # ------------------------------------------------------------------
        if is_coupon_continuation(text, matches):
            append_pending(text, y)
            continue

        # ------------------------------------------------------------------
        # The final percentage is the published portfolio weight.
        # Earlier percentages remain inside the security name.
        # ------------------------------------------------------------------
        last = matches[-1]
        percentage = float(last.group(1))

        if not 0 <= percentage <= 100:
            raise RuntimeError(
                "Spatial parser found a published percentage outside 0-100%."
            )

        percentage_text = clean_text(last.group(0))
        name_before_weight = clean_text(text[:last.start()])

        # If the same visual line contains a name before the final %, it is a
        # complete row.  Any pending wrapped fragment belongs to that row.
        if name_before_weight:
            fragments = list(pending_name_fragments)
            fragments.append(name_before_weight)

            name = combine_holding_name_fragments(fragments)
            commit(name, percentage, percentage_text)
            reset_pending()
            continue

        # A percentage-only line can be the right-hand weight column of a
        # wrapped holding.  It is accepted ONLY if there is already a pending
        # name group and the weight is immediately adjacent vertically.
        if len(matches) == 1 and pending_name_fragments:
            if pending_last_y is None or abs(y - pending_last_y) > WRAPPED_ROW_MAX_GAP:
                raise RuntimeError(
                    "Spatial parser found a percentage-only row outside the "
                    "same wrapped holding row."
                )

            name = combine_holding_name_fragments(
                pending_name_fragments
            )

            commit(name, percentage, percentage_text)
            reset_pending()
            continue

        # A percentage-only row without an already reconstructed holding is
        # not accepted. This prevents chart-axis percentages and unrelated
        # right-column values from becoming holdings.
        raise RuntimeError(
            "Spatial parser found a percentage-only row without a "
            "reconstructed holding name."
        )

    if pending_name_fragments:
        raise RuntimeError(
            "Spatial parser found holding-name text without its published "
            "portfolio weight."
        )

    if not holdings:
        raise RuntimeError(
            "Spatial parser found the Top 10 holdings area but could not "
            "extract any holding/percentage pairs."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            "Spatial parser produced more than 10 holdings."
        )

    ranks = [item["rank"] for item in holdings]
    expected = list(range(1, len(holdings) + 1))
    if ranks != expected:
        raise RuntimeError(
            "Spatial parser holding ranks are not sequential."
        )

    for holding in holdings:
        if not holding["name"]:
            raise RuntimeError(
                "Spatial parser produced an empty holding name."
            )

    return holdings

def extract_holdings_spatial_fallback(
    pdf_bytes: bytes,
) -> tuple[list[dict], str]:
    """
    Recover Top Holdings from the physical PDF layout.

    This fallback scans every page containing a visible "Top 10 Holdings"
    heading, reconstructs the left visual column, and parses the physical rows
    directly. It does NOT pass the reconstructed text through the ordinary
    rank-based parser because PDF object order can detach names from weights.
    """

    positioned = _positioned_pdf_lines(
        pdf_bytes
    )

    if not positioned:
        raise RuntimeError(
            "Spatial PDF fallback could not extract positioned text."
        )

    candidates = []

    pages = sorted({
        line["page"]
        for line in positioned
    })

    for page_number in pages:
        page_lines = [
            line
            for line in positioned
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
            if _is_spatial_top_holdings_heading(
                line["text"]
            )
        ]

        for heading in headings:
            section = _build_spatial_holdings_section(
                page_lines,
                heading,
                page_width,
            )

            if not clean_text(section):
                continue

            # Recover the exact visual rows again so the spatial parser can
            # use y/x evidence rather than the flattened text representation.
            left_edge = max(
                0.0,
                heading["x0"] - 15.0,
            )

            right_edge = _spatial_column_boundary(
                page_lines,
                heading,
                page_width,
            )

            candidate_rows = []

            for line in page_lines:
                if line["y"] >= heading["y"] - 2.0:
                    continue

                if line["y"] < heading["y"] - 520.0:
                    continue

                fragments = []

                for fragment in line.get("fragments", []):
                    x0 = float(
                        fragment.get(
                            "x0",
                            0.0,
                        )
                    )

                    if x0 < left_edge:
                        continue

                    if x0 >= right_edge:
                        continue

                    fragments.append(
                        fragment
                    )

                if not fragments:
                    continue

                fragments.sort(
                    key=lambda item: item["x0"]
                )

                row_text = clean_text(
                    " ".join(
                        fragment["text"]
                        for fragment in fragments
                    )
                )

                if not row_text:
                    continue

                if _is_spatial_top_holdings_heading(
                    row_text
                ):
                    break

                if is_holdings_end(
                    row_text
                ):
                    break

                candidate_rows.append(
                    {
                        "page": page_number,
                        "y": line["y"],
                        "x0": min(
                            fragment["x0"]
                            for fragment in fragments
                        ),
                        "x1": max(
                            fragment["x0"]
                            for fragment in fragments
                        ),
                        "text": row_text,
                        "fragments": fragments,
                    }
                )

            candidate_rows.sort(
                key=lambda item: -item["y"]
            )

            try:
                holdings = _parse_spatial_holdings_rows(
                    candidate_rows
                )
            except Exception as error:
                candidates.append(
                    {
                        "page": page_number,
                        "section": section,
                        "holdings": None,
                        "error": clean_text(
                            str(error)
                        ),
                    }
                )
                continue

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

    valid.sort(
        key=lambda candidate: (
            len(candidate["holdings"]),
            sum(
                len(item["name"])
                for item in candidate["holdings"]
            ),
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
# NEXT-STAGE FIXED-INCOME RECOVERY
# =============================================================================


def parse_holdings_fixed_income_recovery(
    section_text: str,
) -> list[dict]:
    """
    Conservative fixed-income recovery parser.

    This stage runs ONLY after the complete Recovery 2 engine has failed a
    fund.  It handles two official-PDF text layouts without using coordinate
    proximity pairing:

    1. Row-oriented layout:
         rank -> security description -> published weight

    2. Split-column text layout:
         ranked security descriptions first, followed by the published
         portfolio-weight column in the same published order.

    Fixed-income security descriptions may contain coupon/rate percentages.
    Earlier percentages remain part of the security name.  A portfolio weight
    is selected only when the text structure provides an unambiguous published
    weight for that holding.

    No coordinate proximity pairing, fuzzy matching, estimation,
    interpolation, fabrication, or third-party data is used.
    """

    lines = build_logical_holding_lines(
        section_text
    )

    if not lines:
        raise RuntimeError(
            "Fixed-income recovery parser received an empty section."
        )

    def extract_fixed_rank(
        line: str,
    ) -> tuple[int | None, str]:
        """
        Extract a genuine published holding rank.

        This deliberately does NOT treat date fragments such as
        1-JUL-2039 as holding rank 1.
        """

        text = clean_text(line)
        if not text:
            return None, ""

        match = re.match(
            r"^(\d{1,2})(?:[.)]|:)[ \t]+(.+)$",
            text,
        )

        if match:
            rank = int(match.group(1))
            if 1 <= rank <= MAX_HOLDINGS:
                return rank, clean_text(match.group(2))

        match = re.match(
            r"^(\d{1,2})[ \t]+(.+)$",
            text,
        )

        if match:
            rank = int(match.group(1))
            remainder = clean_text(match.group(2))
            if (
                1 <= rank <= MAX_HOLDINGS
                and not re.match(
                    r"^[A-Za-z]{3,9}-\d{4}\b",
                    remainder,
                )
                and not re.match(
                    r"^[A-Za-z]{3,9}-",
                    remainder,
                )
            ):
                return rank, remainder

        if re.fullmatch(r"(?:10|[1-9])", text):
            return int(text), ""

        return None, text

    def append_name_fragment(
        fragments: list[str],
        value: str,
    ) -> None:
        fragment = clean_holding_fragment(value)
        if fragment:
            fragments.append(fragment)

    def build_ranked_blocks() -> list[dict]:
        """
        Build textual blocks beginning at explicit published ranks.

        Blocks are purely text-order constructs.  No x/y coordinate is used.
        """

        blocks: list[dict] = []
        current: dict | None = None

        for raw_line in lines:
            line = clean_text(raw_line)
            if not line:
                continue

            if is_holding_header_or_noise(line):
                continue

            rank, remainder = extract_fixed_rank(line)

            if rank is not None:
                if current is not None:
                    blocks.append(current)

                current = {
                    "rank": rank,
                    "lines": [],
                }

                if remainder:
                    current["lines"].append(remainder)

                continue

            if current is not None:
                current["lines"].append(line)

        if current is not None:
            blocks.append(current)

        return blocks

    def validate_rank_sequence(
        blocks: list[dict],
    ) -> None:
        ranks = [int(block["rank"]) for block in blocks]
        if not ranks:
            raise RuntimeError(
                "Fixed-income recovery parser found no published holding ranks."
            )

        expected = list(range(1, len(ranks) + 1))
        if ranks != expected:
            raise RuntimeError(
                "Fixed-income recovery published ranks are not sequential. "
                f"Parsed={ranks}; Expected={expected}"
            )

    def parse_row_oriented_blocks(
        blocks: list[dict],
    ) -> tuple[list[dict], int]:
        """
        Parse a normal row-oriented representation.

        A portfolio weight is accepted only when it is structurally the final
        percentage of the holding block.  If a block contains only one
        percentage, that percentage must be on a standalone line; otherwise
        it is too ambiguous to distinguish a coupon from the portfolio weight.
        """

        holdings: list[dict] = []
        multiple_percentage_blocks = 0

        for block in blocks:
            block_lines = [
                clean_text(line)
                for line in block["lines"]
                if clean_text(line)
                and not is_holding_header_or_noise(line)
            ]

            if not block_lines:
                raise RuntimeError(
                    f"Rank {block['rank']} has no security description text."
                )

            all_matches: list[tuple[float, str, int, int, str]] = []

            for line in block_lines:
                for match in percentage_matches(line):
                    value = float(match.group(1))
                    if 0 <= value <= 100:
                        all_matches.append(
                            (
                                value,
                                clean_text(match.group(0)),
                                match.start(),
                                match.end(),
                                line,
                            )
                        )

            if not all_matches:
                raise RuntimeError(
                    f"Rank {block['rank']} has no published percentage in its "
                    "text block."
                )

            if len(all_matches) > 1:
                multiple_percentage_blocks += 1

            selected = all_matches[-1]
            weight, weight_text, start, end, selected_line = selected

            selected_line_index = len(block_lines) - 1
            for index in range(len(block_lines) - 1, -1, -1):
                if block_lines[index] == selected_line:
                    selected_line_index = index
                    break

            is_trailing_percentage = (
                selected_line_index == len(block_lines) - 1
                and end == len(selected_line)
            )

            if not is_trailing_percentage:
                raise RuntimeError(
                    f"Rank {block['rank']} has a percentage that is not the "
                    "final token of its holding block."
                )

            if len(all_matches) == 1:
                # A lone inline percentage can be a coupon and cannot be
                # distinguished safely from a portfolio weight.  Only a
                # standalone final percentage is unambiguous in this case.
                if clean_text(selected_line) != weight_text:
                    raise RuntimeError(
                        f"Rank {block['rank']} has only one percentage and it "
                        "is not a standalone published weight."
                    )

            fragments: list[str] = []
            selected_consumed = False

            for line in block_lines:
                if not selected_consumed and line == selected_line:
                    matches = percentage_matches(line)
                    last = matches[-1] if matches else None
                    if (
                        last is not None
                        and last.start() == start
                        and last.end() == end
                    ):
                        prefix = clean_text(line[:start])
                        if prefix:
                            append_name_fragment(
                                fragments,
                                prefix,
                            )
                        selected_consumed = True
                        continue

                append_name_fragment(
                    fragments,
                    line,
                )

            name = combine_holding_name_fragments(
                fragments
            )

            if not name:
                raise RuntimeError(
                    f"Rank {block['rank']} has a published percentage but no "
                    "holding name."
                )

            holdings.append(
                {
                    "rank": int(block["rank"]),
                    "name": name,
                    "weightPercent": weight,
                    "weightText": weight_text,
                }
            )

        return holdings, multiple_percentage_blocks

    def parse_split_column_layout(
        blocks: list[dict],
    ) -> tuple[list[dict], int]:
        """
        Parse a text extraction where ranked security descriptions are emitted
        first and the portfolio-weight column is emitted after the final
        ranked description.

        The only accepted pairing is published text order:

            rank/name 1 ... rank/name N -> weight 1 ... weight N

        No coordinate proximity is used.  The parser requires an exact count
        match between ranked names and trailing published weights.
        """

        count = len(blocks)
        if count < 1:
            raise RuntimeError("No ranked security blocks found.")

        names: list[str] = []

        for block_index, block in enumerate(blocks):
            block_lines = [
                clean_text(line)
                for line in block["lines"]
                if clean_text(line)
                and not is_holding_header_or_noise(line)
            ]

            if not block_lines:
                raise RuntimeError(
                    f"Rank {block['rank']} has no security description text."
                )

            # For every block except the final one, all text belongs to the
            # security description because the trailing weight column cannot
            # begin until the final ranked security has been emitted.
            if block_index < count - 1:
                fragments = []
                for line in block_lines:
                    append_name_fragment(fragments, line)
                name = combine_holding_name_fragments(fragments)
                if not name:
                    raise RuntimeError(
                        f"Rank {block['rank']} has an empty security description."
                    )
                names.append(name)
                continue

            # The final rank block may contain the beginning of a separate
            # weight column.  Locate a trailing run of standalone percentage
            # lines.  These are structurally different from coupon text such
            # as "2.375% 1-JUL-2039" because the whole line is the percentage.
            weight_lines: list[str] = []
            name_lines = list(block_lines)

            while name_lines:
                candidate = name_lines[-1]
                matches = percentage_matches(candidate)
                if not matches:
                    break

                if any(
                    match.start() != 0
                    or match.end() != len(candidate)
                    for match in matches
                ):
                    break

                name_lines.pop()
                weight_lines.insert(0, candidate)

            # Also support one trailing line containing the complete weight
            # column, e.g. "5.8% 5.1% 4.9% ...".
            if not weight_lines and name_lines:
                candidate = name_lines[-1]
                matches = percentage_matches(candidate)
                if (
                    len(matches) == count
                    and all(
                        0 <= float(match.group(1)) <= 100
                        for match in matches
                    )
                ):
                    weight_lines = [candidate]
                    name_lines.pop()

            if not weight_lines:
                raise RuntimeError(
                    "Split-column structure did not expose a trailing "
                    "published weight column."
                )

            trailing_weights: list[tuple[float, str]] = []
            for weight_line in weight_lines:
                matches = percentage_matches(weight_line)
                for match in matches:
                    value = float(match.group(1))
                    if not 0 <= value <= 100:
                        raise RuntimeError(
                            "Split-column published weight is outside 0-100%."
                        )
                    trailing_weights.append(
                        (
                            value,
                            clean_text(match.group(0)),
                        )
                    )

            if len(trailing_weights) != count:
                raise RuntimeError(
                    "Split-column fixed-income structure was not established: "
                    f"found {len(trailing_weights)} trailing published "
                    f"weights for {count} ranked holdings."
                )

            # The final security name is everything remaining before the
            # trailing weight column.  Earlier coupon percentages are kept.
            fragments = []
            for line in name_lines:
                append_name_fragment(fragments, line)
            final_name = combine_holding_name_fragments(fragments)
            if not final_name:
                raise RuntimeError(
                    f"Rank {block['rank']} has an empty security description."
                )
            names.append(final_name)

            holdings = []
            for index, block_for_holding in enumerate(blocks):
                weight, weight_text = trailing_weights[index]
                name = clean_holding_name(names[index])
                if not name:
                    raise RuntimeError(
                        f"Rank {block_for_holding['rank']} has an empty "
                        "security description."
                    )
                holdings.append(
                    {
                        "rank": int(block_for_holding["rank"]),
                        "name": name,
                        "weightPercent": weight,
                        "weightText": weight_text,
                    }
                )

            coupon_evidence = sum(
                len(percentage_matches(name))
                for name in names
            )

            if coupon_evidence == 0:
                raise RuntimeError(
                    "Split-column structure was established, but no coupon "
                    "percentage evidence was found in the security names."
                )

            return holdings, coupon_evidence

        raise RuntimeError(
            "Split-column parser could not construct published holdings."
        )

    blocks = build_ranked_blocks()

    # The first ten published ranks are the only valid target universe for
    # this stage.  Anything after rank 10 is not silently consumed as a new
    # holding universe.
    if len(blocks) > MAX_HOLDINGS:
        blocks = blocks[:MAX_HOLDINGS]

    validate_rank_sequence(blocks)

    row_error = None
    try:
        holdings, multiple_percentage_blocks = parse_row_oriented_blocks(
            blocks
        )
    except Exception as error:
        row_error = clean_text(str(error))
        holdings = []
        multiple_percentage_blocks = 0

    if not holdings:
        try:
            holdings, split_evidence = parse_split_column_layout(
                blocks
            )
            multiple_percentage_blocks = max(
                multiple_percentage_blocks,
                split_evidence,
            )
        except Exception as split_error:
            raise RuntimeError(
                "Fixed-income parser could not establish an unambiguous "
                "ranked holding/weight structure. "
                f"Row-oriented={row_error or 'not attempted'}; "
                f"Split-column={clean_text(str(split_error))}"
            ) from split_error

    if not holdings:
        raise RuntimeError(
            "Fixed-income parser found no published holdings."
        )

    if multiple_percentage_blocks == 0:
        raise RuntimeError(
            "Fixed-income recovery evidence not found: no coupon/portfolio "
            "percentage ambiguity was detected."
        )

    holdings = validate_holdings(
        holdings,
        "Fixed-income recovery",
    )

    ranks = [
        item["rank"]
        for item in holdings
    ]

    expected_ranks = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    if ranks != expected_ranks:
        raise RuntimeError(
            "Fixed-income recovery published ranks are not sequential. "
            f"Parsed={ranks}; Expected={expected_ranks}"
        )

    return holdings


def extract_fixed_income_recovery_engine(
    pdf_bytes: bytes,
) -> dict:
    """
    Run the next specialised recovery engine against an official Prudential PDF.

    Stage order:

        1. Existing strict fixed-income text parser.
        2. Physical-PDF spatial recovery using the successful row-59/60
           extraction architecture.

    The spatial stage uses the physical PDF only to reconstruct genuine visual
    rows. It never pairs an arbitrary name with an arbitrary nearby weight.
    A name and weight must be present on the same reconstructed visual row;
    otherwise the candidate is rejected. Fixed-income rows use the LAST % as
    the published portfolio weight, preserving earlier coupon % values inside
    the security name.
    """

    full_text, page_count = extract_pdf_text(pdf_bytes)
    section_text, section_status = extract_holdings_section(full_text)

    fixed_income_error = None

    # -------------------------------------------------------------------------
    # Stage 1: existing strict fixed-income parser
    # -------------------------------------------------------------------------
    if section_status != "not_published" and clean_text(section_text):
        try:
            holdings = parse_holdings_fixed_income_recovery(section_text)

            return {
                "status": "success",
                "holdings": holdings,
                "parser": "fixed_income_table",
                "fixedIncomeParserError": None,
                "spatialRecoveryError": None,
                "fullText": full_text,
                "sectionText": section_text,
                "pageCount": page_count,
            }
        except Exception as error:
            fixed_income_error = clean_text(str(error))
    else:
        fixed_income_error = (
            "Normal text extraction did not expose a published Top Holdings "
            "section."
        )

    # -------------------------------------------------------------------------
    # Stage 2: physical-PDF visual-row recovery
    # -------------------------------------------------------------------------
    spatial_error = None

    try:
        spatial_holdings, spatial_section = extract_holdings_spatial_fallback(
            pdf_bytes
        )

        spatial_holdings = validate_holdings(
            spatial_holdings,
            "Fixed-income spatial recovery",
        )

        ranks = [int(item["rank"]) for item in spatial_holdings]
        expected_ranks = list(range(1, len(spatial_holdings) + 1))
        if ranks != expected_ranks:
            raise RuntimeError(
                "Fixed-income spatial recovery published ranks are not "
                f"sequential. Parsed={ranks}; Expected={expected_ranks}"
            )

        # Require fixed-income evidence. This specialised stage must not become
        # a generic rescue parser for unrelated failed funds.
        coupon_evidence = sum(
            len(percentage_matches(item["name"]))
            for item in spatial_holdings
        )
        if coupon_evidence == 0:
            raise RuntimeError(
                "Fixed-income spatial recovery found no coupon/rate "
                "percentage evidence in the published security names."
            )

        return {
            "status": "success",
            "holdings": spatial_holdings,
            "parser": "fixed_income_spatial_fallback",
            "fixedIncomeParserError": fixed_income_error,
            "spatialRecoveryError": None,
            "fullText": full_text,
            "sectionText": spatial_section,
            "pageCount": page_count,
        }

    except Exception as error:
        spatial_error = clean_text(str(error))

    raise HoldingsParseFailure(
        "Fixed-income recovery failed after both the text parser and the "
        "physical-PDF spatial recovery. "
        f"FixedIncome={fixed_income_error}; Spatial={spatial_error}",
        section_text=section_text or "",
        full_text=full_text,
        diagnostics={
            "fixedIncomeParserError": fixed_income_error,
            "spatialRecoveryError": spatial_error,
        },
    )

def recover_single_fund_fixed_income(
    page,
    excel_fund: dict,
) -> dict:
    """
    Re-download the official Prudential factsheet for one fund and run ONLY
    the specialised fixed-income recovery parser.
    """

    excel_row = int(
        excel_fund["excelRow"]
    )

    prudential_url = ensure_prudential_url(
        excel_fund["prudentialUrl"]
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

    final_url = clean_text(page.url)

    if not is_prudential_url(final_url):
        raise RuntimeError(
            "Fixed-income recovery page redirected outside Prudential "
            f"Singapore: {final_url}"
        )

    page_title = clean_text(page.title())
    fund_name = extract_fund_page_name(page)
    factsheet_url = find_factsheet_url(page)

    print(
        f"Fixed-income recovery row {excel_row}: "
        f"{fund_name or '-'}"
    )

    factsheet_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    engine = extract_fixed_income_recovery_engine(
        factsheet_bytes
    )

    if engine["status"] == "no_holdings_section":
        return {
            "status": "no_holdings_section",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "finalUrl": final_url,
            "pageTitle": page_title,
            "fundName": fund_name,
            "excelPruAccessName": excel_fund.get("pruAccessName"),
            "factsheetUrl": factsheet_url,
            "factsheetPageCount": engine["pageCount"],
            "factsheetDocumentDate": extract_document_date(
                engine["fullText"]
            ),
            "factsheetDataAsAt": extract_data_as_at(
                engine["fullText"]
            ),
            "topHoldingsCount": 0,
            "topHoldings": [],
            "holdingsParser": None,
            "recoveryEngine": "fixed_income_no_holdings_section",
            "recoveryStage": "recovery2_next",
            "_factsheetBytes": factsheet_bytes,
            "_fullText": engine["fullText"],
            "_sectionText": "",
        }

    holdings = engine["holdings"]

    return {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pageTitle": page_title,
        "fundName": fund_name,
        "excelPruAccessName": excel_fund.get("pruAccessName"),
        "factsheetUrl": factsheet_url,
        "factsheetDocumentDate": extract_document_date(
            engine["fullText"]
        ),
        "factsheetDataAsAt": extract_data_as_at(
            engine["fullText"]
        ),
        "factsheetPageCount": engine["pageCount"],
        "holdingsSectionStatus": "published",
        "topHoldingsCount": len(holdings),
        "topHoldings": holdings,
        "holdingsParser": engine["parser"],
        "recoveryEngine": "fixed_income_table",
        "recoveryStage": "recovery2_next",
        "primaryParserError": None,
        "fallbackParserError": None,
        "spatialRecoveryError": None,
        "spatialDiagnostics": None,
        "fixedIncomeParserError": None,
        "_factsheetBytes": factsheet_bytes,
        "_fullText": engine["fullText"],
        "_sectionText": engine["sectionText"],
    }




# =============================================================================
# EXTRACTION ENGINE
# =============================================================================

def extract_with_recovery_engine(
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

        return {
            "status":
                "no_holdings_section",

            "holdings":
                [],

            "parser":
                None,

            "primaryParserError":
                None,

            "fallbackParserError":
                None,

            "spatialRecoveryError":
                None,

            "fullText":
                full_text,

            "sectionText":
                "",

            "pageCount":
                page_count,

            "spatialDiagnostics":
                None,
        }

    primary_error = None

    try:

        holdings = parse_holdings(
            section_text
        )

        return {
            "status":
                "success",

            "holdings":
                holdings,

            "parser":
                "primary",

            "primaryParserError":
                None,

            "fallbackParserError":
                None,

            "spatialRecoveryError":
                None,

            "fullText":
                full_text,

            "sectionText":
                section_text,

            "pageCount":
                page_count,

            "spatialDiagnostics":
                None,
        }

    except Exception as error:

        primary_error = clean_text(
            str(error)
        )

    fallback_error = None

    try:

        holdings = parse_holdings_fallback(
            section_text
        )

        return {
            "status":
                "success",

            "holdings":
                holdings,

            "parser":
                "fallback",

            "primaryParserError":
                primary_error,

            "fallbackParserError":
                None,

            "spatialRecoveryError":
                None,

            "fullText":
                full_text,

            "sectionText":
                section_text,

            "pageCount":
                page_count,

            "spatialDiagnostics":
                None,
        }

    except Exception as error:

        fallback_error = clean_text(
            str(error)
        )

    spatial_error = None

    try:

        (
            spatial_holdings,
            spatial_section,
        ) = extract_holdings_spatial_fallback(
            pdf_bytes
        )

        spatial_diagnostics = {
            "parser": "spatial_fallback",
            "parsedHoldingCount": len(spatial_holdings),
        }

        return {
            "status":
                "success",

            "holdings":
                spatial_holdings,

            "parser":
                "spatial_fallback",

            "primaryParserError":
                primary_error,

            "fallbackParserError":
                fallback_error,

            "spatialRecoveryError":
                None,

            "fullText":
                full_text,

            "sectionText":
                spatial_section,

            "pageCount":
                page_count,

            "spatialDiagnostics":
                spatial_diagnostics,
        }

    except Exception as error:

        spatial_error = clean_text(
            str(error)
        )

    raise HoldingsParseFailure(
        "All recovery extraction engines failed. "
        f"Primary={primary_error}; "
        f"Fallback={fallback_error}; "
        f"Spatial={spatial_error}",
        section_text=section_text,
        full_text=full_text,
        diagnostics={
            "primaryParserError":
                primary_error,

            "fallbackParserError":
                fallback_error,

            "spatialRecoveryError":
                spatial_error,
        },
    )


# =============================================================================
# FUND PAGE
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
        f"RECOVERY 2 - FUND ROW {excel_row}"
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

    engine = extract_with_recovery_engine(
        factsheet_bytes
    )

    if engine[
        "status"
    ] == "no_holdings_section":

        return {
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

            "factsheetDocumentDate":
                extract_document_date(
                    engine[
                        "fullText"
                    ]
                ),

            "factsheetDataAsAt":
                extract_data_as_at(
                    engine[
                        "fullText"
                    ]
                ),

            "topHoldingsCount":
                0,

            "topHoldings":
                [],

            "holdingsParser":
                None,

            "recoveryEngine":
                "no_holdings_section",

            "_factsheetBytes":
                factsheet_bytes,

            "_fullText":
                engine[
                    "fullText"
                ],

            "_sectionText":
                "",
        }

    holdings = engine[
        "holdings"
    ]

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
            excel_fund.get(
                "pruAccessName"
            ),

        "factsheetUrl":
            factsheet_url,

        "factsheetDocumentDate":
            extract_document_date(
                engine[
                    "fullText"
                ]
            ),

        "factsheetDataAsAt":
            extract_data_as_at(
                engine[
                    "fullText"
                ]
            ),

        "factsheetPageCount":
            engine[
                "pageCount"
            ],

        "holdingsSectionStatus":
            "published",

        "topHoldingsCount":
            len(
                holdings
            ),

        "topHoldings":
            holdings,

        "holdingsParser":
            engine[
                "parser"
            ],

        "primaryParserError":
            engine[
                "primaryParserError"
            ],

        "fallbackParserError":
            engine[
                "fallbackParserError"
            ],

        "spatialRecoveryError":
            engine[
                "spatialRecoveryError"
            ],

        "spatialDiagnostics":
            engine[
                "spatialDiagnostics"
            ],

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

    return result


# =============================================================================
# EXACT FINAL VERIFICATION
# =============================================================================

def verify_exact_against_official_pdf(
    page,
    result: dict,
    extraction_engine=None,
) -> dict:

    print(
        "FINAL VERIFICATION: downloading official PDF again..."
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
            "Final verification factsheet download returned "
            f"HTTP {response.status}."
        )

    verification_bytes = response.body()

    if not verification_bytes.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Final verification response was not a PDF."
        )

    if extraction_engine is None:
        extraction_engine = extract_with_recovery_engine

    verification_engine = extraction_engine(
        verification_bytes
    )

    if verification_engine[
        "status"
    ] != "success":

        raise RuntimeError(
            "Final verification could not extract published holdings."
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

        diagnostics = {
            "originalSignature":
                original_signature,

            "verifiedSignature":
                verified_signature,

            "originalHoldings":
                original_holdings,

            "verifiedHoldings":
                verified_holdings,

            "verificationParser":
                verification_engine[
                    "parser"
                ],

            "verificationPrimaryParserError":
                verification_engine[
                    "primaryParserError"
                ],

            "verificationFallbackParserError":
                verification_engine[
                    "fallbackParserError"
                ],

            "verificationSpatialDiagnostics":
                verification_engine.get(
                    "spatialDiagnostics"
                ),

            "verificationFixedIncomeParserError":
                verification_engine.get(
                    "fixedIncomeParserError"
                ),
        }

        raise HoldingsParseFailure(
            "EXACT VERIFICATION FAILED: the re-downloaded "
            "official Prudential PDF produced a different "
            "rank/name/weight signature.",
            section_text=verification_engine[
                "sectionText"
            ],
            full_text=verification_engine[
                "fullText"
            ],
            diagnostics=diagnostics,
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
# OUTPUT HELPERS
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
        urlparse(
            result.get(
                "finalUrl"
            )
            or
            result.get(
                "prudentialUrl"
            )
        ).path.rstrip(
            "/"
        ).split(
            "/"
        )[-1]
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

    verification_bytes = (
        verification[
            "verificationFactsheetBytes"
        ]
    )

    verification_full_text = (
        verification[
            "verificationFullText"
        ]
    )

    verification_section = (
        verification[
            "verificationSectionText"
        ]
    )

    (
        directory
        /
        "factsheet.pdf"
    ).write_bytes(
        verification_bytes
    )

    (
        directory
        /
        "factsheet_text.txt"
    ).write_text(
        verification_full_text,
        encoding="utf-8",
    )

    (
        directory
        /
        "top_holdings_section.txt"
    ).write_text(
        verification_section,
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
            "recoveryEngine":
                result.get(
                    "recoveryEngine"
                ),

            "holdingsParser":
                result.get(
                    "holdingsParser"
                ),

            "primaryParserError":
                result.get(
                    "primaryParserError"
                ),

            "fallbackParserError":
                result.get(
                    "fallbackParserError"
                ),

            "spatialRecoveryError":
                result.get(
                    "spatialRecoveryError"
                ),

            "spatialDiagnostics":
                result.get(
                    "spatialDiagnostics"
                ),

            "verification":
                {
                    "verifiedParser":
                        verification[
                            "verifiedParser"
                        ],

                    "exactSignatureMatch":
                        True,
                },

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

            "holdingsParser":
                result.get(
                    "holdingsParser"
                ),

            "verificationParser":
                verification.get(
                    "verifiedParser"
                ),

            "exactVerification":
                True,

            "savedAtUtc":
                utc_now_iso(),
        },
    )

    return directory


def save_no_holdings_section(
    result: dict,
) -> Path:

    directory = recovery_directory(
        result
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    clean_result = {
        key:
            value
        for key, value
        in result.items()
        if not key.startswith("_")
    }

    clean_result[
        "recoveryEngine"
    ] = "no_holdings_section"

    save_json(
        directory
        /
        "top_holdings.json",
        clean_result,
    )

    (
        directory
        /
        "factsheet_text.txt"
    ).write_text(
        result.get(
            "_fullText"
        )
        or
        "",
        encoding="utf-8",
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

            "status":
                "no_holdings_section",

            "savedAtUtc":
                utc_now_iso(),
        },
    )

    return directory


def save_failure(
    excel_fund: dict,
    error_text: str,
    attempt_diagnostics: list[dict],
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

        "attempts":
            attempt_diagnostics,

        "failedAtUtc":
            utc_now_iso(),
    }

    save_json(
        directory
        /
        "failure.json",
        failure,
    )

    if section_text:

        (
            directory
            /
            "top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    if full_text:

        (
            directory
            /
            "factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if diagnostics:

        save_json(
            directory
            /
            "recovery_diagnostics.json",
            diagnostics,
        )

    return directory


# =============================================================================
# NEXT RECOVERY FAILURE OUTPUT
# =============================================================================

def save_next_recovery_failure(
    excel_fund: dict,
    error_text: str,
    attempt_diagnostics: list[dict],
    section_text: str | None = None,
    full_text: str | None = None,
    diagnostics: dict | None = None,
) -> Path:
    """Save next-stage diagnostics without destroying the Recovery 2 result."""

    excel_row = int(excel_fund["excelRow"])

    directory = (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure = {
        "status": "failed",
        "recoveryStage": "recovery2_next",
        "recoveryMethod": (
                    final_next_result.get("holdingsParser")
                    if final_next_result
                    else "fixed_income_table"
                ),
        "excelRow": excel_row,
        "prudentialUrl": excel_fund["prudentialUrl"],
        "excelPruAccessName": excel_fund.get("pruAccessName"),
        "error": clean_text(error_text),
        "attempts": attempt_diagnostics,
        "failedAtUtc": utc_now_iso(),
    }

    save_json(
        directory / "recovery2_next_failure.json",
        failure,
    )

    if section_text:
        (
            directory / "recovery2_next_top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    if full_text:
        (
            directory / "recovery2_next_factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if diagnostics:
        save_json(
            directory / "recovery2_next_diagnostics.json",
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
        + "#" * 72
    )

    print(
        "VGRAT FMS - PRUDENTIAL TOP HOLDINGS RECOVERY 2"
    )

    print(
        "#" * 72
    )

    print(
        f"Started UTC: {started_at}"
    )

    # -------------------------------------------------------------------------
    # Load Recovery 1 source
    # -------------------------------------------------------------------------

    recovery_1 = load_recovery_1_run_summary()

    print(
        "\nRECOVERY 1 SOURCE"
    )

    print(
        f"Recovery 1 status: "
        f"{recovery_1.get('status')}"
    )

    print(
        f"Recovery 1 recorded universe: "
        f"{recovery_1.get('excelFundUniverse')}"
    )

    print(
        f"Recovery 1 still-failed funds: "
        f"{recovery_1.get('failedFunds')}"
    )

    # -------------------------------------------------------------------------
    # Load Excel
    # -------------------------------------------------------------------------

    excel_funds = read_excel_funds()

    recovery_universe = build_recovery_universe(
        recovery_1,
        excel_funds,
    )

    print(
        f"\nRecovery universe: "
        f"{len(recovery_universe)} failed fund(s)"
    )

    if not recovery_universe:

        completed_at = utc_now_iso()

        summary = {
            "status":
                "nothing_to_recover",

            "startedAtUtc":
                started_at,

            "completedAtUtc":
                completed_at,

            "recovery1RunSummary":
                str(
                    RECOVERY_1_RUN_SUMMARY_FILE
                ),

            "excelFile":
                str(
                    EXCEL_FILE
                ),

            "recovery1StillFailedFunds":
                int(
                    recovery_1.get(
                        "failedFunds",
                        0,
                    )
                    or
                    0
                ),

            "recoveryUniverse":
                0,

            "recoveredFunds":
                0,

            "noHoldingsSectionFunds":
                0,

            "failedFunds":
                0,

            "rules": {
                "recovery1StillFailedFundsOnly":
                    True,

                "officialPrudentialOnly":
                    True,

                "noThirdPartyHoldings":
                    True,

                "noInferredHoldings":
                    True,

                "noFabricatedPercentages":
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
            "\nNothing to recover."
        )

        return 0

    # -------------------------------------------------------------------------
    # Recovery state
    # -------------------------------------------------------------------------

    recovered = []

    no_holdings_section = []

    failed = []

    fallback_recovered = []

    spatial_recovered = []

    # Results from the automatic post-Recovery-2 specialised stage.
    next_recovery_attempted = []
    next_recovery_recovered = []
    next_recovery_failed = []

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
            f"RECOVERY FUND {index}/{len(recovery_universe)}"
        )

        print(
            f"Excel row: {excel_row}"
        )

        print(
            "=" * 78
        )

        attempt_diagnostics = []

        final_result = None

        last_exception = None

        for attempt in range(
            1,
            RETRY_COUNT + 1,
        ):

            print(
                f"\nAttempt {attempt}/{RETRY_COUNT}"
            )

            attempt_record = {
                "attempt":
                    attempt,

                "startedAtUtc":
                    utc_now_iso(),
            }

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

                        if (
                            result[
                                "status"
                            ]
                            ==
                            "no_holdings_section"
                        ):

                            directory = (
                                save_no_holdings_section(
                                    result
                                )
                            )

                            result[
                                "outputDirectory"
                            ] = str(
                                directory
                            )

                            final_result = result

                        else:

                            # -------------------------------------------------
                            # Exact final verification
                            # -------------------------------------------------

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

                if final_result:

                    attempt_record[
                        "parser"
                    ] = final_result.get(
                        "holdingsParser"
                    )

                    attempt_record[
                        "status"
                    ] = final_result.get(
                        "status"
                    )

                attempt_diagnostics.append(
                    attempt_record
                )

                last_exception = None

                break

            except Exception as error:

                last_exception = error

                error_text = clean_text(
                    str(error)
                )

                attempt_record[
                    "status"
                ] = "failed"

                attempt_record[
                    "error"
                ] = error_text

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

                attempt_diagnostics.append(
                    attempt_record
                )

                print(
                    f"Attempt failed: {error_text}"
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
        # No successful recovery
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
                    "Unknown recovery failure."
                ),
                attempt_diagnostics,
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
                            "Unknown recovery failure."
                        ),

                    "outputDirectory":
                        str(
                            directory
                        ),

                    "attempts":
                        attempt_diagnostics,
                }
            )

            continue

        # ---------------------------------------------------------------------
        # No holdings section
        # ---------------------------------------------------------------------

        if (
            final_result[
                "status"
            ]
            ==
            "no_holdings_section"
        ):

            no_holdings_section.append(
                final_result
            )

            print(
                "\nRECOVERY RESULT: "
                "NO TOP HOLDINGS SECTION"
            )

            continue

        # ---------------------------------------------------------------------
        # Successful recovery
        # ---------------------------------------------------------------------

        recovered.append(
            final_result
        )

        parser = final_result.get(
            "holdingsParser"
        )

        if parser == "fallback":

            fallback_recovered.append(
                final_result
            )

        elif parser == "spatial_fallback":

            spatial_recovered.append(
                final_result
            )

        print(
            "\nRECOVERY SUCCESS"
        )

        print(
            f"Fund: "
            f"{final_result.get('fundName') or '-'}"
        )

        print(
            f"Parser: "
            f"{parser or '-'}"
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
    # AUTOMATIC NEXT RECOVERY STAGE
    # =========================================================================
    #
    # IMPORTANT:
    # This list is created ONLY after the complete Recovery 2 engine has
    # finished processing and exact-verifying every Recovery 2 fund.
    # Nothing is hardcoded by Excel row number.
    #
    # The next specialised parser is therefore isolated to funds that are
    # genuinely still failed at this point.
    # =========================================================================

    recovery2_initial_failed = list(failed)

    next_recovery_rows = {
        int(item["excelRow"])
        for item in recovery2_initial_failed
    }

    next_recovery_universe = [
        excel_funds[row]
        for row in sorted(next_recovery_rows)
        if row in excel_funds
    ]

    print(
        "\n"
        + "=" * 78
    )

    print(
        "AUTOMATIC NEXT RECOVERY STAGE"
    )

    print(
        "=" * 78
    )

    print(
        "This stage is running only against funds that remained FAILED "
        "after the complete Recovery 2 engine."
    )

    print(
        f"Next-stage universe: {len(next_recovery_universe)}"
    )

    for index, excel_fund in enumerate(
        next_recovery_universe,
        start=1,
    ):
        excel_row = int(excel_fund["excelRow"])

        next_recovery_attempted.append(excel_row)

        print(
            "\n"
            + "-" * 78
        )

        print(
            f"NEXT RECOVERY {index}/{len(next_recovery_universe)} - "
            f"Excel row {excel_row}"
        )

        print(
            "Method: fixed_income_table"
        )

        attempt_diagnostics = []
        final_next_result = None
        last_next_exception = None

        for attempt in range(
            1,
            RETRY_COUNT + 1,
        ):
            print(
                f"\nNext-stage attempt {attempt}/{RETRY_COUNT}"
            )

            attempt_record = {
                "attempt": attempt,
                "startedAtUtc": utc_now_iso(),
            }

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
                        result = recover_single_fund_fixed_income(
                            page,
                            excel_fund,
                        )

                        if result["status"] != "success":
                            raise RuntimeError(
                                "Next-stage fixed-income recovery did not "
                                "produce a published Top Holdings result."
                            )

                        verification = verify_exact_against_official_pdf(
                            page,
                            result,
                            extraction_engine=extract_fixed_income_recovery_engine,
                        )

                        directory = save_success(
                            result,
                            verification,
                        )

                        result["outputDirectory"] = str(directory)
                        result["exactVerification"] = True
                        result["verificationParser"] = verification[
                            "verifiedParser"
                        ]
                        result["recoveryStage"] = "recovery2_next"
                        result["recoveryMethod"] = result.get("holdingsParser") or "fixed_income_table"

                        final_next_result = result

                    finally:
                        context.close()
                        browser.close()

                attempt_record["status"] = "success"
                attempt_record["completedAtUtc"] = utc_now_iso()
                attempt_record["parser"] = (
                    final_next_result.get("holdingsParser")
                    if final_next_result
                    else None
                )
                attempt_diagnostics.append(attempt_record)
                last_next_exception = None
                break

            except Exception as error:
                last_next_exception = error

                attempt_record["status"] = "failed"
                attempt_record["error"] = clean_text(str(error))
                attempt_record["completedAtUtc"] = utc_now_iso()

                if isinstance(error, HoldingsParseFailure):
                    attempt_record["diagnostics"] = error.diagnostics

                attempt_diagnostics.append(attempt_record)

                print(
                    f"Next-stage attempt failed: {clean_text(str(error))}"
                )

                if attempt < RETRY_COUNT:
                    print(
                        f"Retrying in {RETRY_DELAY_SECONDS} seconds..."
                    )
                    time.sleep(RETRY_DELAY_SECONDS)

        if final_next_result is not None:
            next_recovery_recovered.append(final_next_result)

            # Replace the earlier Recovery 2 failure with the verified
            # next-stage success in the final state.
            failed = [
                item
                for item in failed
                if int(item["excelRow"]) != excel_row
            ]

            recovered.append(final_next_result)

            print(
                "\nNEXT RECOVERY SUCCESS"
            )

            print(
                f"Fund: {final_next_result.get('fundName') or '-'}"
            )

            print(
                f"Holdings: {final_next_result.get('topHoldingsCount')}"
            )

            print(
                "Exact fixed-income verification: PASS"
            )

            for holding in final_next_result["topHoldings"]:
                print(
                    f"  {holding['rank']}. "
                    f"{holding['name']} - "
                    f"{holding['weightText']}"
                )

            continue

        initial_failure = next(
            (
                item
                for item in recovery2_initial_failed
                if int(item["excelRow"]) == excel_row
            ),
            {},
        )

        section_text = getattr(
            last_next_exception,
            "section_text",
            None,
        )

        full_text = getattr(
            last_next_exception,
            "full_text",
            None,
        )

        diagnostics = getattr(
            last_next_exception,
            "diagnostics",
            None,
        )

        next_failure_directory = save_next_recovery_failure(
            excel_fund,
            clean_text(
                str(last_next_exception)
                if last_next_exception
                else "Unknown next-stage recovery failure."
            ),
            attempt_diagnostics,
            section_text=section_text,
            full_text=full_text,
            diagnostics=diagnostics,
        )

        next_recovery_failed.append(
            {
                "excelRow": excel_row,
                "prudentialUrl": excel_fund["prudentialUrl"],
                "pruAccessName": excel_fund.get("pruAccessName"),
                "error": clean_text(
                    str(last_next_exception)
                    if last_next_exception
                    else "Unknown next-stage recovery failure."
                ),
                "recoveryStage": "recovery2_next",
                "recoveryMethod": (
                    final_next_result.get("holdingsParser")
                    if final_next_result
                    else "fixed_income_table"
                ),
                "initialRecovery2Failure": initial_failure,
                "outputDirectory": str(next_failure_directory),
                "attempts": attempt_diagnostics,
            }
        )

    # Keep the final failed records explicitly tied to their latest recovery
    # stage while preserving the original Recovery 2 failure diagnostics.
    final_failed_by_row = {
        int(item["excelRow"]): item
        for item in failed
    }

    for item in next_recovery_failed:
        final_failed_by_row[int(item["excelRow"])] = item

    failed = [
        final_failed_by_row[row]
        for row in sorted(final_failed_by_row)
    ]

    # =========================================================================
    # CONSOLIDATED SUMMARY
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

        "recovery1RunSummary":
            str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "recovery1ExcelFundUniverse":
            recovery_1.get(
                "excelFundUniverse"
            ),

        "recovery1StillFailedFunds":
            recovery_1.get(
                "failedFunds"
            ),

        "recoveryUniverse":
            len(
                recovery_universe
            ),

        "recoveredFunds":
            len(
                recovered
            ),

        "noHoldingsSectionFunds":
            len(
                no_holdings_section
            ),

        "failedFunds":
            len(
                failed
            ),

        "recovery2InitialFailedFunds":
            len(
                recovery2_initial_failed
            ),

        "recovery2InitialRecoveredFunds":
            len(
                recovered
            ) - len(
                next_recovery_recovered
            ),

        "nextRecoveryAttemptedFunds":
            len(
                next_recovery_attempted
            ),

        "nextRecoveryRecoveredFunds":
            len(
                next_recovery_recovered
            ),

        "nextRecoveryFailedFunds":
            len(
                next_recovery_failed
            ),

        "finalFailedFunds":
            len(
                failed
            ),

        "fallbackRecoveredFunds":
            len(
                fallback_recovered
            ),

        "spatialRecoveredFunds":
            len(
                spatial_recovered
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

                    "holdingsParser":
                        result.get(
                            "holdingsParser"
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

        "fallbackRecoveredFundsDetail":
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

                    "topHoldingsCount":
                        result.get(
                            "topHoldingsCount"
                        ),

                    "primaryParserError":
                        result.get(
                            "primaryParserError"
                        ),
                }

                for result
                in fallback_recovered
            ],

        "spatialRecoveredFundsDetail":
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

                    "topHoldingsCount":
                        result.get(
                            "topHoldingsCount"
                        ),

                    "fallbackParserError":
                        result.get(
                            "fallbackParserError"
                        ),

                    "spatialDiagnostics":
                        result.get(
                            "spatialDiagnostics"
                        ),
                }

                for result
                in spatial_recovered
            ],

        "noHoldingsSectionDetail":
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
                }

                for result
                in no_holdings_section
            ],

        "failedFundsDetail":
            failed,

        "nextRecoveryAttemptedRows":
            next_recovery_attempted,

        "nextRecoveryRecoveredFundsDetail":
            [
                {
                    "excelRow": result.get("excelRow"),
                    "fundName": result.get("fundName"),
                    "topHoldingsCount": result.get("topHoldingsCount"),
                    "holdingsParser": result.get("holdingsParser"),
                    "recoveryStage": result.get("recoveryStage"),
                    "recoveryMethod": result.get("recoveryMethod"),
                    "exactVerification": result.get("exactVerification"),
                    "verificationParser": result.get("verificationParser"),
                    "outputDirectory": result.get("outputDirectory"),
                }
                for result in next_recovery_recovered
            ],

        "nextRecoveryFailedFundsDetail":
            next_recovery_failed,

        "recovery1FailureField":
            "stillFailedFundsDetail",

        "recovery1FailedRows":
            [
                int(item["excelRow"])
                for item in recovery_universe
            ],

        "rules":
            {
                "recovery1StillFailedFundsOnly":
                    True,

                "excelColumnAControlsMasterUniverseForSelectedRows":
                    True,

                "officialPrudentialSingaporeOnly":
                    True,

                "officialFactsheetOnly":
                    True,

                "thirdPartyHoldings":
                    False,

                "noInferredHoldings":
                    True,

                "noFabricatedHoldings":
                    True,

                "noFabricatedPercentages":
                    True,

                "noForcedTenEntries":
                    True,

                "fewerThanTenPublishedHoldingsAllowed":
                    True,

                "duplicateHoldingNamesAllowed":
                    True,

                "duplicateHoldingPercentagesAllowed":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "hyphenatedLineWrapReconstruction":
                    True,

                "fixedIncomeCouponPercentagesSupported":
                    True,

                "primaryParser":
                    True,

                "fallbackParser":
                    True,

                "spatialRecovery":
                    True,

                "spatialRecoveryUsesFallbackParser":
                    True,

                "noCoordinateProximityPairing":
                    True,

                "noFuzzyHoldingMatching":
                    True,

                "noSyntheticValues":
                    True,

                "exactFinalPdfVerification":
                    True,

                "exactRankNameWeightSignatureRequired":
                    True,

                "automaticNextRecoveryAfterRecovery2Failure":
                    True,

                "nextRecoveryUsesOnlyPostRecovery2Failures":
                    True,

                "nextRecoveryHardcodedRows":
                    False,

                "fixedIncomeLastPercentageIsPortfolioWeight":
                    True,

                "fixedIncomeEarlierPercentagesRemainInSecurityName":
                    True,

                "fixedIncomeRecoveryUsesOfficialPdfOnly":
                    True,

                "fixedIncomeRecoveryUsesNoCoordinateProximityPairing":
                    True,

                "fixedIncomeRecoveryRequiresMultiplePercentageEvidence":
                    True,

                "fixedIncomeRecoveryExactSecondPdfVerification":
                    True,
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
        "PRUDENTIAL TOP HOLDINGS RECOVERY 2 COMPLETE"
    )

    print(
        "=" * 78
    )

    print(
        f"Recovery 1 recorded universe: "
        f"{recovery_1.get('excelFundUniverse')}"
    )

    print(
        f"Recovery 1 still-failed funds: "
        f"{recovery_1.get('failedFunds')}"
    )

    print(
        f"Recovery universe: "
        f"{len(recovery_universe)}"
    )

    print(
        f"Recovered: "
        f"{len(recovered)}"
    )

    print(
        f"No Top Holdings section: "
        f"{len(no_holdings_section)}"
    )

    print(
        f"Still failed: "
        f"{len(failed)}"
    )

    print(
        f"Recovery 2 initial failed: "
        f"{len(recovery2_initial_failed)}"
    )

    print(
        f"Automatic next recovery attempted: "
        f"{len(next_recovery_attempted)}"
    )

    print(
        f"Automatic next recovery recovered: "
        f"{len(next_recovery_recovered)}"
    )

    print(
        f"Automatic next recovery still failed: "
        f"{len(next_recovery_failed)}"
    )

    print(
        f"Recovered by fallback parser: "
        f"{len(fallback_recovered)}"
    )

    print(
        f"Recovered by spatial fallback: "
        f"{len(spatial_recovered)}"
    )

    print(
        f"Total published holdings recovered: "
        f"{total_holdings}"
    )

    print(
        "\nRecovery order:"
    )

    print(
        "  1. Normal PDF extraction"
    )

    print(
        "  2. Primary parser"
    )

    print(
        "  3. Fallback parser"
    )

    print(
        "  4. Spatial section reconstruction"
    )

    print(
        "  5. Fallback parser on reconstructed section"
    )

    print(
        "  6. Exact official PDF re-download"
    )

    print(
        "  7. Exact rank/name/weight verification"
    )

    print(
        "  8. Automatic next-stage fixed-income recovery "
        "for remaining failures only"
    )

    print(
        "  9. Exact fixed-income re-download and "
        "rank/name/weight verification"
    )

    print(
        "\nNo fuzzy matching."
    )

    print(
        "No proximity name/weight pairing."
    )

    print(
        "No fabricated holdings."
    )

    print(
        "No fabricated percentages."
    )

    print(
        "No third-party holdings."
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
            "\nSTILL FAILED:"
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

    # Deliberately return zero even when some recovery funds remain failed.
    # GitHub Actions should preserve and upload diagnostics instead of
    # discarding the recovery artifacts because of a non-zero process exit.
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
            "\nFATAL RECOVERY ERROR:",
            file=sys.stderr,
        )

        print(
            str(error),
            file=sys.stderr,
        )

        raise SystemExit(
            1
        )
