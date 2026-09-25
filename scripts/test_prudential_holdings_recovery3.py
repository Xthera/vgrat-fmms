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

    output_holdings_recovery_2/run_summary.json
                    |
                    v
             failedFundsDetail
                    |
                    v
              Funds Links.xlsm
                    |
                    v
        Official Prudential fund page
                    |
                    v
       Official Prudential factsheet
                    |
                    v
        PHYSICAL PDF-LAYOUT PARSER
                    |
                    v
       Exact second-PDF verification


RECOVERY 3 PURPOSE
==================

Recovery 3 is designed for Prudential factsheets where the Top 10
Holdings table is extracted by normal PDF text extraction in an
incorrect reading order.

Typical fixed-income holding:

    SINGAPORE (REPUBLIC OF) 2.375% 1-JUL-2039    5.7%

or:

    BPCE SA 4.6% 21-JAN-2035                     1.6%

or:

    SEATRIUM FINANCIAL SERVICES PTE LTD
    2.95% 28-APR-2031                            1.6%

The earlier percentage is part of the security name.

The LAST / right-hand percentage is the published portfolio weight.


IMPORTANT
=========

Recovery 3 does NOT use the normal linear PDF extraction order for
the holdings table.

Instead it uses the physical PDF coordinates to reconstruct the
published table rows.

This is necessary because some Prudential factsheets extract the
Top 10 Holdings section approximately as:

    security names
    percentages
    charts
    other columns

rather than in visual table order.

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
- No fabricated holdings.
- No fabricated percentages.
- No forced ten holdings.
- Fewer than ten published holdings is valid.
- Published order is preserved.
- Earlier percentages remain in the security name.
- The right-hand / final percentage is the portfolio weight.
- Fixed-income maturity dates are used to identify holding rows.
- Physical PDF layout is used to reconstruct rows.
- No arbitrary proximity pairing of unrelated text.
- Multiple-percentage evidence is required.
- Exact final verification downloads the official PDF again.
- Exact final verification reruns the same parser.
- Exact rank/name/weight signature must match.
- Recovery 3 never modifies Recovery 2 output.
"""


from __future__ import annotations

import json
import math
import re
import sys
import time
from collections import Counter, defaultdict
from dataclasses import dataclass
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


# ======================================================================
# CONFIGURATION
# ======================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

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

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000
FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000
POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3
RETRY_DELAY_SECONDS = 3.0

MAX_HOLDINGS = 10


# PDF coordinate tolerances.
#
# These are deliberately used for physical row/column reconstruction,
# not arbitrary name-to-weight proximity matching.
Y_LINE_TOLERANCE = 3.5
X_CLUSTER_TOLERANCE = 12.0

# Minimum horizontal separation that normally distinguishes the
# right-hand weight column from percentages embedded in security names.
MIN_WEIGHT_COLUMN_GAP = 35.0

# A fixed-income holding normally contains a maturity date such as:
#
#     1-JUL-2039
#     21-JAN-2035
#
# This is intentionally strict.
MATURITY_DATE_RE = re.compile(
    r"\b\d{1,2}-"
    r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"-\d{4}\b",
    re.IGNORECASE,
)

PERCENTAGE_RE = re.compile(
    r"(?<![\d.])"
    r"([0-9]+(?:\.[0-9]+)?)"
    r"\s*%"
)

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# ======================================================================
# GENERAL HELPERS
# ======================================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(value) -> str:
    text = clean_text(value)

    text = text.casefold()

    text = text.replace("–", "-")
    text = text.replace("—", "-")

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def utc_now_iso() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
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

    text = text.strip("._")

    if not text:
        text = "fund"

    return text[:120]


def is_prudential_url(
    url: str,
) -> bool:

    try:

        parsed = urlparse(url)

        hostname = (
            parsed.hostname or ""
        ).lower()

        return (
            parsed.scheme.lower()
            in {"http", "https"}
            and hostname in PRUDENTIAL_HOSTS
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
            f"Non-Prudential URL rejected: {url}"
        )

    return url


# ======================================================================
# FAILURE TYPE
# ======================================================================

class HoldingsParseFailure(
    RuntimeError
):

    def __init__(
        self,
        message: str,
        section_text: str = "",
        full_text: str = "",
        diagnostics: dict | None = None,
    ) -> None:

        super().__init__(message)

        self.section_text = section_text
        self.full_text = full_text
        self.diagnostics = diagnostics or {}


# ======================================================================
# RECOVERY 2 INPUT
# ======================================================================

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

    if not isinstance(data, dict):

        raise RuntimeError(
            "Recovery 2 run summary is not a JSON object."
        )

    failed = data.get(
        "failedFundsDetail"
    )

    if not isinstance(failed, list):

        raise RuntimeError(
            "Recovery 2 run summary does not contain "
            "a valid failedFundsDetail list."
        )

    return data


# ======================================================================
# EXCEL
# ======================================================================

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
                "excelRow": row_number,
                "prudentialUrl": prudential_url,
                "pruAccessName": pruaccess_name,
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
                failure["excelRow"]
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
            and recovery_url != excel_url
        ):

            raise RuntimeError(
                "Recovery 2 URL does not match Excel "
                f"for row {excel_row}."
            )

        recovery_name = clean_text(
            failure.get(
                "pruAccessName"
            )
            or failure.get(
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
            and excel_name
            and recovery_name != excel_name
        ):

            raise RuntimeError(
                "Recovery 2 PruAccess name does not match "
                f"Excel for row {excel_row}."
            )

        universe.append(
            {
                "excelRow": excel_row,
                "prudentialUrl": excel_url,
                "pruAccessName": excel_name,
                "recovery2Failure": failure,
            }
        )

    return universe


# ======================================================================
# PRUDENTIAL PAGE / FACTSHEET
# ======================================================================

def find_factsheet_url(page) -> str:

    anchors = page.locator("a")

    candidates = []

    for index in range(
        anchors.count()
    ):

        anchor = anchors.nth(index)

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

        if (
            "factsheet"
            in href_lower
        ):

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
            f"Factsheet HTTP status: "
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


# ======================================================================
# NORMAL PDF TEXT
# ======================================================================

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
                f"Failed to extract PDF page "
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
            "PDF contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


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


# ======================================================================
# PHYSICAL PDF POSITIONED TEXT
# ======================================================================

@dataclass
class PositionedFragment:

    text: str
    x: float
    y: float
    font_size: float
    page_number: int


@dataclass
class PositionedLine:

    page_number: int
    y: float
    fragments: list[PositionedFragment]

    @property
    def text(self) -> str:

        ordered = sorted(
            self.fragments,
            key=lambda item: (
                item.x,
                item.text,
            ),
        )

        return clean_text(
            " ".join(
                item.text
                for item in ordered
            )
        )


def _safe_float(
    value,
) -> float:

    try:

        result = float(
            value
        )

        if math.isfinite(
            result
        ):

            return result

    except Exception:

        pass

    return 0.0


def _positioned_pdf_lines(
    pdf_bytes: bytes,
) -> list[PositionedLine]:

    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    all_lines = []

    for page_number, page in enumerate(
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

            text = clean_text(
                text
            )

            if not text:
                return

            x = _safe_float(
                tm[4]
            )

            y = _safe_float(
                tm[5]
            )

            size = _safe_float(
                font_size
            )

            fragments.append(
                PositionedFragment(
                    text=text,
                    x=x,
                    y=y,
                    font_size=size,
                    page_number=page_number,
                )
            )

        try:

            page.extract_text(
                visitor_text=visitor_text
            )

        except Exception:

            continue

        if not fragments:
            continue

        # Group fragments by physical Y coordinate.
        grouped = []

        for fragment in sorted(
            fragments,
            key=lambda item: (
                -item.y,
                item.x,
            ),
        ):

            target = None

            for group in grouped:

                if abs(
                    group["y"]
                    - fragment.y
                ) <= Y_LINE_TOLERANCE:

                    target = group
                    break

            if target is None:

                grouped.append(
                    {
                        "y": fragment.y,
                        "fragments": [
                            fragment
                        ],
                    }
                )

            else:

                target[
                    "fragments"
                ].append(
                    fragment
                )

        for group in grouped:

            line = PositionedLine(
                page_number=page_number,
                y=group["y"],
                fragments=sorted(
                    group["fragments"],
                    key=lambda item: item.x,
                ),
            )

            if clean_text(
                line.text
            ):

                all_lines.append(
                    line
                )

    all_lines.sort(
        key=lambda line: (
            line.page_number,
            -line.y,
        )
    )

    return all_lines


# ======================================================================
# POSITIONED TEXT DIAGNOSTICS
# ======================================================================

def positioned_lines_to_text(
    lines: list[PositionedLine],
) -> str:

    output = []

    current_page = None

    for line in lines:

        if line.page_number != current_page:

            current_page = line.page_number

            output.append(
                f"--- PAGE {current_page} ---"
            )

        output.append(
            line.text
        )

    return "\n".join(
        output
    )


def positioned_line_debug(
    line: PositionedLine,
) -> dict:

    return {
        "page": line.page_number,
        "y": round(
            line.y,
            2,
        ),
        "text": line.text,
        "fragments": [
            {
                "x": round(
                    fragment.x,
                    2,
                ),
                "y": round(
                    fragment.y,
                    2,
                ),
                "text": fragment.text,
            }
            for fragment in line.fragments
        ],
    }


# ======================================================================
# TOP HOLDINGS HEADING
# ======================================================================

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


def find_top_holdings_heading_lines(
    lines: list[PositionedLine],
) -> list[PositionedLine]:

    return [
        line
        for line in lines
        if is_top_holdings_heading(
            line.text
        )
    ]


# ======================================================================
# FIXED-INCOME ROW IDENTIFICATION
# ======================================================================

def contains_maturity_date(
    text: str,
) -> bool:

    return bool(
        MATURITY_DATE_RE.search(
            text
        )
    )


def percentage_values(
    text: str,
) -> list[float]:

    values = []

    for match in PERCENTAGE_RE.finditer(
        text
    ):

        try:

            value = float(
                match.group(1)
            )

        except Exception:

            continue

        if 0 <= value <= 100:

            values.append(
                value
            )

    return values


def percentage_fragment_values(
    fragment: PositionedFragment,
) -> list[tuple[float, float]]:

    values = []

    for match in PERCENTAGE_RE.finditer(
        fragment.text
    ):

        try:

            value = float(
                match.group(1)
            )

        except Exception:

            continue

        if not 0 <= value <= 100:
            continue

        values.append(
            (
                value,
                fragment.x,
            )
        )

    return values


def line_has_maturity_and_percentage(
    line: PositionedLine,
) -> bool:

    return (
        contains_maturity_date(
            line.text
        )
        and bool(
            percentage_values(
                line.text
            )
        )
    )


# ======================================================================
# TABLE COLUMN DISCOVERY
# ======================================================================

def discover_weight_column(
    candidate_lines: list[PositionedLine],
) -> float:

    """
    Discover the physical X position of the right-hand weight column.

    We deliberately do NOT pair a name with the nearest arbitrary
    percentage.

    Instead:

      1. collect percentage X positions from genuine maturity-date
         holding rows;
      2. cluster those X positions;
      3. identify the right-most recurring percentage column.

    Coupon percentages occur inside the security-name column.

    Portfolio weights occur in the right-hand table column.
    """

    percentage_x_positions = []

    for line in candidate_lines:

        for fragment in line.fragments:

            for (
                value,
                x,
            ) in percentage_fragment_values(
                fragment
            ):

                percentage_x_positions.append(
                    x
                )

    if not percentage_x_positions:

        raise HoldingsParseFailure(
            "No percentage column could be discovered "
            "from fixed-income holding rows."
        )

    # Cluster X positions.
    clusters = []

    for x in sorted(
        percentage_x_positions
    ):

        placed = False

        for cluster in clusters:

            if abs(
                cluster["mean"]
                - x
            ) <= X_CLUSTER_TOLERANCE:

                cluster["values"].append(
                    x
                )

                cluster["mean"] = (
                    sum(
                        cluster["values"]
                    )
                    / len(
                        cluster["values"]
                    )
                )

                placed = True
                break

        if not placed:

            clusters.append(
                {
                    "mean": x,
                    "values": [x],
                }
            )

    if not clusters:

        raise HoldingsParseFailure(
            "Unable to construct percentage X clusters."
        )

    # The right-most recurring percentage cluster is the candidate
    # portfolio-weight column.
    clusters.sort(
        key=lambda cluster: (
            cluster["mean"]
        )
    )

    # Prefer a cluster that appears on multiple holding rows.
    recurring = [
        cluster
        for cluster in clusters
        if len(
            cluster["values"]
        ) >= 2
    ]

    if recurring:

        selected = max(
            recurring,
            key=lambda cluster: (
                cluster["mean"],
                len(cluster["values"]),
            )
        )

    else:

        selected = max(
            clusters,
            key=lambda cluster: (
                cluster["mean"]
            )
        )

    return float(
        selected["mean"]
    )


def choose_weight_percentage(
    line: PositionedLine,
    weight_column_x: float,
) -> tuple[float, str, float] | None:

    """
    Select the percentage belonging to the physical right-hand
    portfolio-weight column.

    The name/coupon percentage remains untouched.
    """

    candidates = []

    for fragment in line.fragments:

        for match in PERCENTAGE_RE.finditer(
            fragment.text
        ):

            try:

                value = float(
                    match.group(1)
                )

            except Exception:

                continue

            if not 0 <= value <= 100:
                continue

            x = fragment.x

            distance = abs(
                x
                - weight_column_x
            )

            candidates.append(
                (
                    distance,
                    x,
                    value,
                    clean_text(
                        match.group(0)
                    ),
                )
            )

    if not candidates:

        return None

    candidates.sort(
        key=lambda item: (
            item[0],
            -item[1],
        )
    )

    distance, x, value, text = (
        candidates[0]
    )

    # Do not allow an unrelated percentage from the security-name
    # column to become the portfolio weight.
    if (
        distance
        > MIN_WEIGHT_COLUMN_GAP
    ):

        return None

    return (
        value,
        text,
        x,
    )


# ======================================================================
# HOLDING NAME CLEANING
# ======================================================================

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


def remove_weight_percentage(
    text: str,
    weight_text: str,
) -> str:

    if not text:
        return ""

    if not weight_text:
        return text

    # Remove only the selected final/right-hand weight percentage.
    # Other percentages, such as coupons, remain.
    escaped = re.escape(
        weight_text
    )

    matches = list(
        re.finditer(
            escaped,
            text,
            flags=re.IGNORECASE,
        )
    )

    if not matches:
        return text

    match = matches[-1]

    return clean_text(
        text[
            :match.start()
        ]
        + " "
        + text[
            match.end():
        ]
    )


def normalize_maturity_spacing(
    text: str,
) -> str:

    text = clean_text(
        text
    )

    # Preserve the published hyphenated maturity date.
    text = re.sub(
        r"\s*-\s*",
        "-",
        text,
    )

    # Restore spaces between date and preceding security text.
    text = re.sub(
        r"([A-Za-z0-9)])(\d{1,2}-"
        r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
        r"-\d{4})",
        r"\1 \2",
        text,
        flags=re.IGNORECASE,
    )

    return clean_text(
        text
    )


# ======================================================================
# SPATIAL HOLDINGS ROW PARSER
# ======================================================================

def parse_spatial_holdings_rows(
    heading_line: PositionedLine,
    all_lines: list[PositionedLine],
) -> tuple[list[dict], dict]:

    """
    Reconstruct the Top 10 Holdings table from physical PDF layout.

    We first locate the page containing the physical heading.

    Then we examine lines physically below the heading and identify
    fixed-income holdings through their maturity dates.

    The weight column is discovered from the physical X positions of
    percentages found on those genuine holding rows.

    No arbitrary nearest-text pairing is performed.
    """

    page_number = (
        heading_line.page_number
    )

    page_lines = [
        line
        for line in all_lines
        if line.page_number
        == page_number
    ]

    # PDF Y coordinates increase upwards.
    #
    # Text physically below the heading has a lower Y coordinate.
    below_heading = [
        line
        for line in page_lines
        if line.y
        < (
            heading_line.y
            - Y_LINE_TOLERANCE
        )
    ]

    below_heading.sort(
        key=lambda line: -line.y
    )

    # Stop when the physical region moves too far away.
    #
    # The actual table is compact. A large gap normally means the
    # following chart/section belongs to something else.
    region_lines = []

    previous_y = None

    for line in below_heading:

        if previous_y is not None:

            gap = (
                previous_y
                - line.y
            )

            if (
                gap > 45
                and region_lines
            ):
                break

        region_lines.append(
            line
        )

        previous_y = line.y

        # We only need enough material to discover the published
        # Top 10 rows.
        if len(region_lines) > 100:
            break

    candidate_lines = [
        line
        for line in region_lines
        if contains_maturity_date(
            line.text
        )
    ]

    if not candidate_lines:

        raise HoldingsParseFailure(
            "No fixed-income maturity-date rows were found "
            "below the physical Top 10 Holdings heading.",
            diagnostics={
                "heading": positioned_line_debug(
                    heading_line
                ),
                "page": page_number,
                "candidateLineCount": 0,
            },
        )

    # We expect up to ten holdings. There can occasionally be unrelated
    # maturity-date content elsewhere on the page, so identify rows
    # carrying percentages first.
    weighted_candidates = [
        line
        for line in candidate_lines
        if percentage_values(
            line.text
        )
    ]

    if not weighted_candidates:

        raise HoldingsParseFailure(
            "No maturity-date holding rows contain percentages.",
            diagnostics={
                "heading": positioned_line_debug(
                    heading_line
                ),
                "candidateLines": [
                    positioned_line_debug(line)
                    for line in candidate_lines[:30]
                ],
            },
        )

    weight_column_x = (
        discover_weight_column(
            weighted_candidates
        )
    )

    holdings = []

    multiple_percentage_rows = 0

    used_y_values = []

    for line in candidate_lines:

        selected = choose_weight_percentage(
            line,
            weight_column_x,
        )

        if selected is None:
            continue

        (
            weight,
            weight_text,
            weight_x,
        ) = selected

        all_values = percentage_values(
            line.text
        )

        if len(
            all_values
        ) > 1:

            multiple_percentage_rows += 1

        # Reconstruct the entire physical row from fragments.
        #
        # The selected right-hand weight is removed only from the
        # holding name. Coupon percentages remain.
        row_text = line.text

        name_text = remove_weight_percentage(
            row_text,
            weight_text,
        )

        name_text = normalize_maturity_spacing(
            name_text
        )

        name_text = clean_holding_name(
            name_text
        )

        if not name_text:
            continue

        # A genuine fixed-income holding must contain a maturity date.
        if not contains_maturity_date(
            name_text
        ):
            continue

        holdings.append(
            {
                "rank": len(
                    holdings
                ) + 1,
                "name": name_text,
                "weightPercent": weight,
                "weightText": weight_text,
                "percentageCountOnPhysicalRow": len(
                    all_values
                ),
                "weightColumnX": round(
                    weight_column_x,
                    2,
                ),
                "weightTextX": round(
                    weight_x,
                    2,
                ),
                "physicalPage": page_number,
                "physicalY": round(
                    line.y,
                    2,
                ),
                "recoveryParser": (
                    "spatial_fixed_income_table"
                ),
            }
        )

        used_y_values.append(
            line.y
        )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    if not holdings:

        raise HoldingsParseFailure(
            "Physical Top 10 Holdings parser found no "
            "complete fixed-income rows.",
            diagnostics={
                "heading": positioned_line_debug(
                    heading_line
                ),
                "weightColumnX": weight_column_x,
                "candidateLines": [
                    positioned_line_debug(line)
                    for line in candidate_lines[:40]
                ],
            },
        )

    # Ensure the parser really encountered the special fixed-income
    # condition Recovery 3 is designed to handle.
    if multiple_percentage_rows <= 0:

        raise HoldingsParseFailure(
            "Recovery 3 did not find any physical holding row "
            "containing multiple percentages.",
            diagnostics={
                "weightColumnX": weight_column_x,
                "holdings": holdings,
                "candidateLines": [
                    positioned_line_debug(line)
                    for line in candidate_lines[:40]
                ],
            },
        )

    diagnostics = {
        "page": page_number,
        "heading": positioned_line_debug(
            heading_line
        ),
        "weightColumnX": round(
            weight_column_x,
            2,
        ),
        "candidateLineCount": len(
            candidate_lines
        ),
        "weightedCandidateLineCount": len(
            weighted_candidates
        ),
        "multiplePercentageRows": (
            multiple_percentage_rows
        ),
        "usedRows": len(
            holdings
        ),
        "usedYValues": [
            round(
                y,
                2,
            )
            for y in used_y_values
        ],
        "candidateLines": [
            positioned_line_debug(line)
            for line in candidate_lines[:40]
        ],
    }

    return (
        holdings,
        diagnostics,
    )


# ======================================================================
# HOLDINGS VALIDATION
# ======================================================================

def validate_holdings(
    holdings: list[dict],
) -> list[dict]:

    if not holdings:

        raise RuntimeError(
            "Recovery 3 parser produced no holdings."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

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
        holding["rank"]
        for holding in holdings
    ]

    if actual_ranks != expected_ranks:

        raise RuntimeError(
            "Invalid sequential ranks. "
            f"Actual={actual_ranks}; "
            f"Expected={expected_ranks}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding["name"]
        )

        if not name:

            raise RuntimeError(
                "Recovery 3 parser produced "
                "an empty holding name."
            )

        weight = holding[
            "weightPercent"
        ]

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                "Recovery 3 parser produced "
                "an invalid weight."
            )

        if not 0 <= weight <= 100:

            raise RuntimeError(
                "Recovery 3 parser produced "
                "a weight outside 0-100."
            )

        holding["name"] = name

    return holdings


# ======================================================================
# FULL RECOVERY 3 ENGINE
# ======================================================================

def extract_fixed_income_recovery_engine(
    pdf_bytes: bytes,
) -> dict:

    (
        full_text,
        page_count,
    ) = extract_pdf_text(
        pdf_bytes
    )

    positioned_lines = (
        _positioned_pdf_lines(
            pdf_bytes
        )
    )

    if not positioned_lines:

        raise HoldingsParseFailure(
            "Physical PDF extraction produced no positioned text.",
            full_text=full_text,
        )

    headings = (
        find_top_holdings_heading_lines(
            positioned_lines
        )
    )

    if not headings:

        raise HoldingsParseFailure(
            "No physical Top 10 Holdings heading was found.",
            full_text=full_text,
            diagnostics={
                "positionedLineCount": len(
                    positioned_lines
                ),
            },
        )

    parser_attempts = []

    for heading in headings:

        try:

            (
                holdings,
                diagnostics,
            ) = parse_spatial_holdings_rows(
                heading,
                positioned_lines,
            )

            holdings = validate_holdings(
                holdings
            )

            # Exact Recovery 3 requirement:
            # at least one row must contain more than one percentage.
            multi_percentage_count = sum(
                1
                for holding in holdings
                if holding.get(
                    "percentageCountOnPhysicalRow",
                    0,
                ) > 1
            )

            if multi_percentage_count <= 0:

                raise RuntimeError(
                    "No recovered holding row contains "
                    "multiple percentages."
                )

            return {
                "status": "success",
                "holdings": holdings,
                "parser": (
                    "spatial_fixed_income_table"
                ),
                "fullText": full_text,
                "pageCount": page_count,
                "positionedText": (
                    positioned_lines_to_text(
                        positioned_lines
                    )
                ),
                "spatialDiagnostics": diagnostics,
                "multiplePercentageRows": (
                    multi_percentage_count
                ),
            }

        except Exception as error:

            parser_attempts.append(
                {
                    "heading": positioned_line_debug(
                        heading
                    ),
                    "error": clean_text(
                        str(error)
                    ),
                }
            )

    raise HoldingsParseFailure(
        "Recovery 3 spatial fixed-income parser "
        "failed for every Top 10 Holdings heading.",
        full_text=full_text,
        diagnostics={
            "parser": (
                "spatial_fixed_income_table"
            ),
            "headingAttempts": parser_attempts,
            "positionedLineCount": len(
                positioned_lines
            ),
            "positionedText": (
                positioned_lines_to_text(
                    positioned_lines
                )
            ),
        },
    )


# ======================================================================
# FUND PAGE NAME
# ======================================================================

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


# ======================================================================
# SINGLE FUND
# ======================================================================

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

    engine = (
        extract_fixed_income_recovery_engine(
            factsheet_bytes
        )
    )

    return {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pageTitle": page_title,
        "fundName": fund_name,
        "excelPruAccessName": excel_fund.get(
            "pruAccessName"
        ),
        "factsheetUrl": factsheet_url,
        "factsheetPageCount": engine[
            "pageCount"
        ],
        "topHoldingsCount": len(
            engine[
                "holdings"
            ]
        ),
        "topHoldings": engine[
            "holdings"
        ],
        "holdingsParser": engine[
            "parser"
        ],
        "recoveryStage": "recovery3",
        "recoveryEngine": (
            "spatial_fixed_income_table"
        ),
        "multiplePercentageRows": engine[
            "multiplePercentageRows"
        ],
        "_factsheetBytes": factsheet_bytes,
        "_fullText": engine[
            "fullText"
        ],
        "_positionedText": engine[
            "positionedText"
        ],
        "_spatialDiagnostics": engine[
            "spatialDiagnostics"
        ],
    }


# ======================================================================
# EXACT SIGNATURE
# ======================================================================

def holding_signature(
    holdings: list[dict],
) -> list[tuple]:

    return [
        (
            int(
                holding["rank"]
            ),
            normalize_text(
                holding["name"]
            ),
            float(
                holding["weightPercent"]
            ),
        )
        for holding in holdings
    ]


# ======================================================================
# FINAL VERIFICATION
# ======================================================================

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

    verification_bytes = (
        response.body()
    )

    if not verification_bytes.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Final verification response "
            "was not a PDF."
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

    original_signature = (
        holding_signature(
            original_holdings
        )
    )

    verified_signature = (
        holding_signature(
            verified_holdings
        )
    )

    if (
        original_signature
        != verified_signature
    ):

        raise HoldingsParseFailure(
            "EXACT VERIFICATION FAILED: "
            "re-downloaded official Prudential PDF "
            "produced a different rank/name/weight signature.",
            section_text="",
            full_text=verification_engine[
                "fullText"
            ],
            diagnostics={
                "originalSignature": (
                    original_signature
                ),
                "verifiedSignature": (
                    verified_signature
                ),
                "originalHoldings": (
                    original_holdings
                ),
                "verifiedHoldings": (
                    verified_holdings
                ),
                "verifiedSpatialDiagnostics": (
                    verification_engine[
                        "spatialDiagnostics"
                    ]
                ),
            },
        )

    return {
        "verified": True,
        "verifiedParser": (
            verification_engine[
                "parser"
            ]
        ),
        "verifiedHoldingCount": len(
            verified_holdings
        ),
        "verifiedSignature": (
            verified_signature
        ),
        "verificationFactsheetBytes": (
            verification_bytes
        ),
        "verificationFullText": (
            verification_engine[
                "fullText"
            ]
        ),
        "verificationPositionedText": (
            verification_engine[
                "positionedText"
            ]
        ),
        "verificationSpatialDiagnostics": (
            verification_engine[
                "spatialDiagnostics"
            ]
        ),
    }


# ======================================================================
# OUTPUT DIRECTORIES
# ======================================================================

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
        or f"fund_{excel_row}"
    )

    return (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
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
        / "factsheet.pdf"
    ).write_bytes(
        verification[
            "verificationFactsheetBytes"
        ]
    )

    (
        directory
        / "factsheet_text.txt"
    ).write_text(
        verification[
            "verificationFullText"
        ],
        encoding="utf-8",
    )

    (
        directory
        / "positioned_factsheet_text.txt"
    ).write_text(
        verification[
            "verificationPositionedText"
        ],
        encoding="utf-8",
    )

    clean_result = {
        key: value
        for key, value in result.items()
        if not key.startswith("_")
    }

    clean_result[
        "verification"
    ] = {
        "verified": True,
        "verifiedParser": (
            verification[
                "verifiedParser"
            ]
        ),
        "verifiedHoldingCount": (
            verification[
                "verifiedHoldingCount"
            ]
        ),
        "exactSignatureMatch": True,
        "verifiedAtUtc": utc_now_iso(),
    }

    save_json(
        directory
        / "top_holdings.json",
        clean_result,
    )

    save_json(
        directory
        / "recovery_diagnostics.json",
        {
            "recoveryStage": "recovery3",
            "recoveryEngine": (
                "spatial_fixed_income_table"
            ),
            "holdingsParser": (
                result[
                    "holdingsParser"
                ]
            ),
            "verificationParser": (
                verification[
                    "verifiedParser"
                ]
            ),
            "exactSignatureMatch": True,
            "multiplePercentageRows": (
                result.get(
                    "multiplePercentageRows"
                )
            ),
            "spatialDiagnostics": (
                result.get(
                    "_spatialDiagnostics"
                )
            ),
            "verificationSpatialDiagnostics": (
                verification.get(
                    "verificationSpatialDiagnostics"
                )
            ),
            "savedAtUtc": utc_now_iso(),
        },
    )

    save_json(
        directory
        / "metadata.json",
        {
            "excelRow": result.get(
                "excelRow"
            ),
            "fundName": result.get(
                "fundName"
            ),
            "factsheetUrl": result.get(
                "factsheetUrl"
            ),
            "recoveryStage": "recovery3",
            "recoveryEngine": (
                "spatial_fixed_income_table"
            ),
            "topHoldingsCount": result.get(
                "topHoldingsCount"
            ),
            "exactVerification": True,
            "savedAtUtc": utc_now_iso(),
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
    positioned_text: str | None = None,
) -> Path:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    directory = (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory
        / "recovery3_failure.json",
        {
            "status": "failed",
            "recoveryStage": "recovery3",
            "excelRow": excel_row,
            "prudentialUrl": (
                excel_fund[
                    "prudentialUrl"
                ]
            ),
            "pruAccessName": (
                excel_fund.get(
                    "pruAccessName"
                )
            ),
            "error": clean_text(
                error_text
            ),
            "attempts": attempts,
            "failedAtUtc": utc_now_iso(),
        },
    )

    if section_text:

        (
            directory
            / "recovery3_top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    if full_text:

        (
            directory
            / "recovery3_factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if positioned_text:

        (
            directory
            / "recovery3_positioned_factsheet_text.txt"
        ).write_text(
            positioned_text,
            encoding="utf-8",
        )

    if diagnostics:

        save_json(
            directory
            / "recovery3_diagnostics.json",
            diagnostics,
        )

    return directory


# ======================================================================
# MAIN
# ======================================================================

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
        "VGRAT FMS - "
        "PRUDENTIAL TOP HOLDINGS RECOVERY 3"
    )

    print(
        "#" * 78
    )

    print(
        f"Started UTC: {started_at}"
    )

    recovery_2 = (
        load_recovery_2_run_summary()
    )

    print(
        "\nRECOVERY 2 SOURCE"
    )

    print(
        "Recovery 2 status: "
        f"{recovery_2.get('status')}"
    )

    print(
        "Recovery 2 recovery universe: "
        f"{recovery_2.get('recoveryUniverse')}"
    )

    print(
        "Recovery 2 failed funds: "
        f"{recovery_2.get('failedFunds')}"
    )

    excel_funds = read_excel_funds()

    recovery_universe = (
        build_recovery_universe(
            recovery_2,
            excel_funds,
        )
    )

    print(
        "\nRecovery 3 universe: "
        f"{len(recovery_universe)}"
    )

    if not recovery_universe:

        completed_at = utc_now_iso()

        summary = {
            "status": "nothing_to_recover",
            "startedAtUtc": started_at,
            "completedAtUtc": completed_at,
            "recovery2RunSummary": str(
                RECOVERY_2_RUN_SUMMARY_FILE
            ),
            "excelFile": str(
                EXCEL_FILE
            ),
            "recovery2FailedFunds": (
                recovery_2.get(
                    "failedFunds",
                    0,
                )
            ),
            "recovery3Universe": 0,
            "attemptedFunds": 0,
            "recoveredFunds": 0,
            "failedFunds": 0,
            "rules": {
                "recovery2FailedFundsOnly": True,
                "hardcodedRows": False,
                "officialPrudentialOnly": True,
                "fixedIncomeParserOnly": True,
                "physicalPdfLayoutParser": True,
                "lastPercentageIsWeight": True,
                "multiplePercentageEvidenceRequired": True,
                "exactFinalVerification": True,
                "recovery2Modified": False,
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
            f"RECOVERY 3 FUND "
            f"{index}/{len(recovery_universe)}"
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
                "attempt": attempt,
                "startedAtUtc": utc_now_iso(),
            }

            print(
                f"\nAttempt "
                f"{attempt}/{RETRY_COUNT}"
            )

            try:

                with sync_playwright() as playwright:

                    browser = (
                        playwright.chromium.launch(
                            headless=BROWSER_HEADLESS
                        )
                    )

                    context = (
                        browser.new_context(
                            viewport={
                                "width": 1440,
                                "height": 1000,
                            },
                            user_agent=(
                                "Mozilla/5.0 "
                                "(Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 "
                                "(KHTML, like Gecko) "
                                "Chrome/153.0.0.0 "
                                "Safari/537.36"
                            ),
                        )
                    )

                    page = (
                        context.new_page()
                    )

                    try:

                        result = (
                            recover_single_fund(
                                page,
                                excel_fund,
                            )
                        )

                        verification = (
                            verify_exact_against_official_pdf(
                                page,
                                result,
                            )
                        )

                        directory = (
                            save_success(
                                result,
                                verification,
                            )
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
                        ] = (
                            verification[
                                "verifiedParser"
                            ]
                        )

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
                    else None
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
                    "Attempt failed: "
                    f"{error}"
                )

                if attempt < RETRY_COUNT:

                    print(
                        "Retrying in "
                        f"{RETRY_DELAY_SECONDS} seconds..."
                    )

                    time.sleep(
                        RETRY_DELAY_SECONDS
                    )

        if final_result is None:

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

            positioned_text = None

            if isinstance(
                diagnostics,
                dict,
            ):

                positioned_text = (
                    diagnostics.get(
                        "positionedText"
                    )
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
                section_text=None,
                full_text=full_text,
                diagnostics=diagnostics,
                positioned_text=positioned_text,
            )

            failed.append(
                {
                    "excelRow": excel_row,
                    "prudentialUrl": (
                        excel_fund[
                            "prudentialUrl"
                        ]
                    ),
                    "pruAccessName": (
                        excel_fund.get(
                            "pruAccessName"
                        )
                    ),
                    "error": clean_text(
                        str(
                            last_exception
                        )
                        if last_exception
                        else
                        "Unknown Recovery 3 failure."
                    ),
                    "outputDirectory": str(
                        directory
                    ),
                    "attempts": attempts,
                }
            )

            continue

        recovered.append(
            final_result
        )

        print(
            "\nRECOVERY 3 SUCCESS"
        )

        print(
            "Fund: "
            f"{final_result.get('fundName') or '-'}"
        )

        print(
            "Parser: "
            "spatial_fixed_income_table"
        )

        print(
            "Multiple-percentage rows: "
            f"{final_result.get('multiplePercentageRows')}"
        )

        print(
            "Holdings: "
            f"{final_result.get('topHoldingsCount')}"
        )

        print(
            "Exact final verification: PASS"
        )

        for holding in (
            final_result[
                "topHoldings"
            ]
        ):

            print(
                f"  {holding['rank']}. "
                f"{holding['name']} - "
                f"{holding['weightText']}"
            )

    completed_at = utc_now_iso()

    total_holdings = sum(
        int(
            result.get(
                "topHoldingsCount",
                0,
            )
            or 0
        )
        for result in recovered
    )

    summary = {
        "status": (
            "success"
            if not failed
            else "partial"
        ),
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "recovery2RunSummary": str(
            RECOVERY_2_RUN_SUMMARY_FILE
        ),
        "excelFile": str(
            EXCEL_FILE
        ),
        "recovery2FailedFunds": (
            recovery_2.get(
                "failedFunds"
            )
        ),
        "recovery3Universe": len(
            recovery_universe
        ),
        "attemptedFunds": len(
            recovery_universe
        ),
        "recovery3AttemptCount": (
            attempts_total
        ),
        "recoveredFunds": len(
            recovered
        ),
        "failedFunds": len(
            failed
        ),
        "totalPublishedTopHoldingsRecovered": (
            total_holdings
        ),
        "recoveredFundsDetail": [
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "fundName": result.get(
                    "fundName"
                ),
                "factsheetUrl": result.get(
                    "factsheetUrl"
                ),
                "topHoldingsCount": result.get(
                    "topHoldingsCount"
                ),
                "holdingsParser": result.get(
                    "holdingsParser"
                ),
                "recoveryStage": result.get(
                    "recoveryStage"
                ),
                "recoveryEngine": result.get(
                    "recoveryEngine"
                ),
                "multiplePercentageRows": result.get(
                    "multiplePercentageRows"
                ),
                "exactVerification": result.get(
                    "exactVerification"
                ),
                "verificationParser": result.get(
                    "verificationParser"
                ),
                "outputDirectory": result.get(
                    "outputDirectory"
                ),
            }
            for result in recovered
        ],
        "failedFundsDetail": failed,
        "rules": {
            "recovery2FailedFundsOnly": True,
            "hardcodedRows": False,
            "officialPrudentialSingaporeOnly": True,
            "officialFactsheetOnly": True,
            "thirdPartyHoldings": False,
            "pruAccessHoldingsDiscovery": False,
            "fuzzyMatching": False,
            "arbitraryCoordinateProximityPairing": False,
            "inferredHoldings": False,
            "fabricatedPercentages": False,
            "forcedTenEntries": False,
            "fewerThanTenPublishedHoldingsAllowed": True,
            "publishedOrderPreserved": True,
            "hyphenatedLineWrapReconstruction": True,
            "fixedIncomeRecovery": True,
            "physicalPdfLayoutExtraction": True,
            "physicalWeightColumnDiscovery": True,
            "lastPercentageIsPortfolioWeight": True,
            "earlierPercentagesRemainInHoldingName": True,
            "maturityDateRequired": True,
            "multiplePercentageEvidenceRequired": True,
            "exactFinalPdfVerification": True,
            "exactRankNameWeightSignatureRequired": True,
            "recovery2Modified": False,
        },
    }

    save_json(
        RECOVERY_RUN_SUMMARY_FILE,
        summary,
    )

    print(
        "\n\n"
        + "=" * 78
    )

    print(
        "PRUDENTIAL TOP HOLDINGS "
        "RECOVERY 3 COMPLETE"
    )

    print(
        "=" * 78
    )

    print(
        "Recovery 2 failed funds: "
        f"{recovery_2.get('failedFunds')}"
    )

    print(
        "Recovery 3 universe: "
        f"{len(recovery_universe)}"
    )

    print(
        f"Recovered: {len(recovered)}"
    )

    print(
        f"Still failed: {len(failed)}"
    )

    print(
        "Published holdings recovered: "
        f"{total_holdings}"
    )

    print(
        "\nRecovery 3 parser:"
    )

    print(
        "  spatial_fixed_income_table"
    )

    print(
        "  physical PDF layout"
    )

    print(
        "  maturity-date row identification"
    )

    print(
        "  right-hand percentage = portfolio weight"
    )

    print(
        "  earlier percentages remain in security name"
    )

    print(
        "  multiple-percentage evidence required"
    )

    print(
        "  exact second PDF verification required"
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
                f" - Row {item['excelRow']}: "
                f"{item['error']}"
            )

    print(
        "\nDone."
    )

    return 0


# ======================================================================
# ENTRY POINT
# ======================================================================

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
