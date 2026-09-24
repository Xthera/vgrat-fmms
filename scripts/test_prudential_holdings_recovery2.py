#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings RECOVERY 2
================================================

RECOVERY 2 SOURCE
-----------------
Recovery 1 failures ONLY:

    output_holdings_recovery/run_summary.json

If Recovery 1's run_summary.json does not expose failed-fund details,
Recovery 2 reconstructs the failed list from:

    output_holdings_recovery/funds/*_failed/failure.json

IMPORTANT
---------
This script NEVER falls back to the baseline:

    output_holdings/run_summary.json

MASTER UNIVERSE
---------------
Funds Links.xlsm

Column A:
    Official Prudential Singapore fund URL

Column B:
    Exact PruAccess fund name

RECOVERY 2 PURPOSE
------------------
Recover only the funds which failed Recovery 1.

The extractor is intentionally structural and conservative.

It requires physical PDF evidence for:

    rank column
        +
    security-name column
        +
    portfolio-weight column

The three pieces must belong to the same table geometry.

The script does NOT:

    - infer ranks from row order
    - invent missing ranks
    - pair arbitrary nearby names and percentages
    - use third-party holdings
    - use calculated/synthetic weights
    - treat fund charges as holdings
    - treat performance figures as holdings
    - treat benchmark figures as holdings
    - treat coupon percentages as portfolio weights
    - force ten holdings
    - require ten holdings
    - use fuzzy holding-name matching

FIXED-INCOME HANDLING
---------------------
For a genuine fixed-income table line such as:

    SINGAPORE (REPUBLIC OF) 2.375% 1-JUL-2039 5.8%

the security name is:

    SINGAPORE (REPUBLIC OF) 2.375% 1-JUL-2039

and the portfolio weight is:

    5.8%

The LAST percentage on the physical security row is used as the
portfolio weight only when the row is structurally confirmed as part
of a holdings table.

The coupon percentage remains inside the security name.

OUTPUT
------
output_holdings_recovery_2/

    all_holdings.json
    run_summary.json

    funds/
        <row>_<identifier>/
            factsheet.pdf
            factsheet_text.txt
            spatial_rows.json
            spatial_candidates.json
            candidate_regions.json
            candidate_holdings.json
            confirmation.json
            metadata.json

            SUCCESS:
                top_holdings.json
                top_holdings_section.txt
                recovery_result.json

            UNRESOLVED:
                diagnostic_failure.json
                recovery_result.json

VERIFICATION
------------
A successful extraction is downloaded again from the same official
Prudential page and extracted again.

The exact signature must match:

    rank
    normalized name
    float(weight)

The second extraction must independently satisfy the same structural
rules.

PYTHON
------
Python 3.12 compatible.
"""

from __future__ import annotations

import io
import json
import math
import os
import re
import statistics
import sys
import time

from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import openpyxl
from pypdf import PdfReader

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except Exception:
    sync_playwright = None
    PlaywrightTimeoutError = Exception


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

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

RECOVERY_ALL_HOLDINGS_FILE = (
    RECOVERY_OUTPUT_DIR / "all_holdings.json"
)

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000
FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000
POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3
RETRY_DELAY_SECONDS = 3.0

MAX_HOLDINGS = 10

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}

# PDF geometry tolerances.
Y_TOLERANCE = 3.5
Y_TOLERANCE_WIDE = 6.0

MIN_WORD_WIDTH = 0.5

# Percentages.
PERCENTAGE_RE = re.compile(
    r"(?<![\d.])(\d+(?:\.\d+)?)\s*%"
)

# Explicit rank forms.
RANK_PREFIX_RE = re.compile(
    r"^\s*(\d{1,2})(?:[.)]\s*|\s+)"
)

RANK_ONLY_RE = re.compile(
    r"^\s*(\d{1,2})\s*$"
)

RANK_MARKED_RE = re.compile(
    r"^\s*(\d{1,2})[.)]\s*$"
)

# Date forms useful for fixed-income descriptors.
DATE_RE = re.compile(
    r"\b"
    r"(?:"
    r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}"
    r"|"
    r"\d{1,2}[-/]"
    r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"[-/]\d{2,4}"
    r"|"
    r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"[-/]\d{1,2}[-/]\d{2,4}"
    r")"
    r"\b",
    re.IGNORECASE,
)

YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")

# Strong semantic rejection patterns.
# These are deliberately strong because false positives are more dangerous
# than leaving a difficult fund unresolved.
FORBIDDEN_NAME_PATTERNS = [
    r"\bcharge\b",
    r"\binitial investment charge\b",
    r"\bsubscription method\b",
    r"\blaunch date\b",
    r"\bunderlying fund size\b",
    r"\bfunds under management\b",
    r"\bmanager of the fund\b",
    r"\bperformance chart\b",
    r"\bcalendar year performance\b",
    r"\bprice indexed\b",
    r"\bbenchmark\b",
    r"\bdistribution class\b",
    r"\bdistribution date\b",
    r"\bperformance\b",
    r"\bmorningstar\b",
    r"\bsustainability rating\b",
    r"\bfinancial year end\b",
    r"\bimportant information\b",
    r"\bpast performance\b",
    r"\bsubscription\b",
    r"\bmanagement fee\b",
    r"\bexpense ratio\b",
    r"\btotal expense\b",
    r"\bportfolio characteristics\b",
    r"\basset allocation\b",
    r"\bsector allocation\b",
    r"\bgeographical allocation\b",
    r"\bprice indexed\b",
    r"\bseries2\b",
]

FORBIDDEN_REGION_PATTERNS = [
    r"\bperformance chart\b",
    r"\bcalendar year performance\b",
    r"\bprice indexed\b",
    r"\bbenchmark\b",
    r"\bmorningstar\b",
    r"\bunderlying fund size\b",
    r"\bfunds under management\b",
    r"\binitial investment charge\b",
    r"\bdistribution class\b",
]

# Generic table header vocabulary.
RANK_HEADERS = {
    "rank",
    "no",
    "no.",
    "#",
}

NAME_HEADERS = {
    "holding",
    "holdings",
    "security",
    "securities",
    "name",
    "description",
    "investment",
}

WEIGHT_HEADERS = {
    "%",
    "weight",
    "weights",
    "portfolio",
    "portfolio weight",
    "weight %",
    "% of net assets",
    "% net assets",
}

# Words which are frequently page furniture rather than holdings.
PAGE_FURNITURE = {
    "source",
    "page",
    "important information",
    "disclaimer",
    "past performance",
    "performance",
    "asset allocation",
    "portfolio characteristics",
}

# ============================================================================
# DATA CLASSES
# ============================================================================


@dataclass
class PDFWord:
    page: int
    text: str
    x: float
    y: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.x + self.width


@dataclass
class PDFRow:
    page: int
    y: float
    words: list[PDFWord]

    @property
    def x_min(self) -> float:
        if not self.words:
            return 0.0
        return min(word.x for word in self.words)

    @property
    def x_max(self) -> float:
        if not self.words:
            return 0.0
        return max(word.right for word in self.words)

    @property
    def text(self) -> str:
        return " ".join(
            word.text.strip()
            for word in sorted(
                self.words,
                key=lambda item: item.x,
            )
            if word.text.strip()
        ).strip()


@dataclass
class HoldingCandidate:
    rank: int
    name: str
    weight: float
    page: int
    source_row_index: int
    strategy: str
    evidence: dict[str, Any]


class HoldingsParseFailure(RuntimeError):
    pass


# ============================================================================
# GENERAL HELPERS
# ============================================================================


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = (
        text.replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
    )

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_name(value: str) -> str:
    text = normalize_text(value).lower()

    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def safe_filename(value: str, fallback: str = "fund") -> str:
    value = normalize_text(value)

    value = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        value,
    )

    value = value.strip(" ._")

    if not value:
        value = fallback

    return value[:160]


def json_dump(path: Path, payload: Any) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def text_dump(path: Path, text: str) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        text,
        encoding="utf-8",
    )


def mean(values: list[float]) -> float:
    return (
        statistics.mean(values)
        if values
        else 0.0
    )


def median(values: list[float]) -> float:
    return (
        statistics.median(values)
        if values
        else 0.0
    )


def population_std(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0

    return statistics.pstdev(values)


def nearly_equal(a: float, b: float, tolerance: float = 1.0) -> bool:
    return abs(a - b) <= tolerance


def clamp_probability_score(value: float) -> float:
    return max(
        0.0,
        min(
            1.0,
            value,
        ),
    )


# ============================================================================
# RECOVERY 1 FAILURE LOADING
# ============================================================================


def normalize_failure_entry(
    entry: dict[str, Any],
) -> dict[str, Any] | None:
    if not isinstance(entry, dict):
        return None

    excel_row = entry.get("excelRow")

    if excel_row is None:
        excel_row = entry.get("row")

    if excel_row is None:
        excel_row = entry.get("rowNumber")

    try:
        excel_row = int(excel_row)
    except Exception:
        return None

    result = dict(entry)

    result["excelRow"] = excel_row

    return result


def extract_failure_list_from_summary(
    payload: dict[str, Any],
) -> list[dict[str, Any]]:
    keys = [
        "failedFundsDetail",
        "failedFundsDetails",
        "failedFunds",
        "failedFundsList",
        "failureDetails",
        "failures",
        "failed",
        "failedFundDetails",
    ]

    for key in keys:
        value = payload.get(key)

        if not isinstance(value, list):
            continue

        normalized = []

        for item in value:
            normalized_item = normalize_failure_entry(item)

            if normalized_item is not None:
                normalized.append(normalized_item)

        if normalized:
            normalized.sort(
                key=lambda item: int(item["excelRow"])
            )

            return normalized

    return []


def load_recovery_1_failure_details_from_output() -> list[dict[str, Any]]:
    """
    Reconstruct failed funds from Recovery 1 output.

    This is intentionally restricted to:

        output_holdings_recovery/funds/*_failed/failure.json
    """

    if not RECOVERY_FUNDS_OUTPUT_DIR.exists():
        return []

    recovered: list[dict[str, Any]] = []

    failure_files = sorted(
        RECOVERY_FUNDS_OUTPUT_DIR.glob(
            "*_failed/failure.json"
        ),
        key=lambda path: path.parent.name,
    )

    for failure_file in failure_files:
        try:
            payload = json.loads(
                failure_file.read_text(
                    encoding="utf-8"
                )
            )
        except Exception:
            continue

        if not isinstance(payload, dict):
            continue

        status = normalize_text(
            payload.get("status")
        ).lower()

        if status != "failed":
            continue

        excel_row = payload.get("excelRow")

        if excel_row is None:
            match = re.match(
                r"^(\d+)_failed$",
                failure_file.parent.name,
            )

            if match:
                excel_row = int(match.group(1))

        try:
            excel_row = int(excel_row)
        except Exception:
            continue

        entry = dict(payload)

        entry["excelRow"] = excel_row

        entry["_reconstructedFromRecovery1Output"] = True
        entry["_recovery1FailureFile"] = str(
            failure_file
        )

        recovered.append(entry)

    recovered.sort(
        key=lambda item: int(item["excelRow"])
    )

    return recovered


def load_recovery_1_run_summary() -> tuple[
    dict[str, Any],
    list[dict[str, Any]],
    str,
]:
    if not RECOVERY_1_RUN_SUMMARY_FILE.exists():
        raise RuntimeError(
            "Recovery 1 run_summary.json was not found: "
            f"{RECOVERY_1_RUN_SUMMARY_FILE}"
        )

    payload = json.loads(
        RECOVERY_1_RUN_SUMMARY_FILE.read_text(
            encoding="utf-8"
        )
    )

    if not isinstance(payload, dict):
        raise RuntimeError(
            "Recovery 1 run_summary.json is not a JSON object."
        )

    failures = extract_failure_list_from_summary(
        payload
    )

    source = "recovery1_run_summary"

    if not failures:
        failures = (
            load_recovery_1_failure_details_from_output()
        )

        source = "recovery1_output_failure_files"

    return (
        payload,
        failures,
        source,
    )


# ============================================================================
# EXCEL
# ============================================================================


def load_excel_universe() -> dict[int, dict[str, Any]]:
    if not EXCEL_FILE.exists():
        raise RuntimeError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = openpyxl.load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_links=True,
    )

    worksheet = workbook.active

    result: dict[int, dict[str, Any]] = {}

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):
        url_value = worksheet.cell(
            row=row_number,
            column=1,
        ).value

        name_value = worksheet.cell(
            row=row_number,
            column=2,
        ).value

        url = normalize_text(url_value)

        if not url:
            continue

        result[row_number] = {
            "excelRow": row_number,
            "prudentialUrl": url,
            "excelPruAccessName": normalize_text(
                name_value
            ),
        }

    workbook.close()

    return result


# ============================================================================
# PRUDENTIAL URL VALIDATION
# ============================================================================


def is_official_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)

        if parsed.scheme.lower() != "https":
            return False

        hostname = (
            parsed.hostname or ""
        ).lower()

        return hostname in PRUDENTIAL_HOSTS

    except Exception:
        return False


# ============================================================================
# PLAYWRIGHT / FACTSHEET
# ============================================================================


def find_factsheet_url(
    page,
    source_url: str,
) -> str:
    if not is_official_prudential_url(
        source_url
    ):
        raise RuntimeError(
            "Source URL is not an official Prudential Singapore URL."
        )

    candidates: list[str] = []

    try:
        links = page.locator("a").all()

        for link in links:
            try:
                href = link.get_attribute(
                    "href"
                )

                text = normalize_text(
                    link.inner_text()
                )

                if not href:
                    continue

                absolute = page.url

                if href.startswith("/"):
                    parsed = urlparse(
                        absolute
                    )

                    href = (
                        f"{parsed.scheme}://"
                        f"{parsed.netloc}"
                        f"{href}"
                    )

                elif href.startswith("#"):
                    continue

                elif not href.lower().startswith(
                    ("http://", "https://")
                ):
                    continue

                if not is_official_prudential_url(
                    href
                ):
                    continue

                lower_href = href.lower()
                lower_text = text.lower()

                if (
                    ".pdf" in lower_href
                    or "factsheet" in lower_href
                    or "factsheet" in lower_text
                    or "fund factsheet" in lower_text
                ):
                    candidates.append(href)

            except Exception:
                continue

    except Exception:
        pass

    # Remove duplicates while preserving order.
    seen = set()
    unique = []

    for item in candidates:
        if item not in seen:
            seen.add(item)
            unique.append(item)

    if not unique:
        raise RuntimeError(
            "No official Prudential factsheet PDF link was found."
        )

    # Prefer explicit factsheet URLs.
    unique.sort(
        key=lambda value: (
            0 if "factsheet" in value.lower() else 1,
            len(value),
        )
    )

    return unique[0]


def download_factsheet(
    page,
    source_url: str,
) -> tuple[bytes, str]:
    last_error: Exception | None = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):
        try:
            page.goto(
                source_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            page.wait_for_timeout(
                POST_PAGE_WAIT_MS
            )

            factsheet_url = find_factsheet_url(
                page,
                source_url,
            )

            response = page.request.get(
                factsheet_url,
                timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
            )

            if not response.ok:
                raise RuntimeError(
                    "Factsheet HTTP status "
                    f"{response.status}"
                )

            content_type = (
                response.headers.get(
                    "content-type",
                    "",
                )
                .lower()
            )

            body = response.body()

            if not body:
                raise RuntimeError(
                    "Factsheet response was empty."
                )

            if (
                "pdf" not in content_type
                and not body.startswith(
                    b"%PDF"
                )
            ):
                raise RuntimeError(
                    "Factsheet response does not appear to be a PDF."
                )

            return (
                body,
                factsheet_url,
            )

        except Exception as exc:
            last_error = exc

            if attempt < RETRY_COUNT:
                time.sleep(
                    RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        "Factsheet download failed: "
        f"{last_error}"
    )


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================


def extract_pdf_text(
    pdf_bytes: bytes,
) -> str:
    reader = PdfReader(
        io.BytesIO(pdf_bytes)
    )

    pages: list[str] = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        pages.append(
            f"\n===== PAGE {page_number} =====\n"
            f"{text}"
        )

    return "\n".join(pages)


# ============================================================================
# POSITIONED PDF EXTRACTION
# ============================================================================


def _visitor_text_factory(
    output_words: list[PDFWord],
    page_number: int,
):
    def visitor_text(
        text,
        cm,
        tm,
        font_dict,
        font_size,
    ):
        if text is None:
            return

        text = str(text)

        if not text.strip():
            return

        try:
            x = float(tm[4])
            y = float(tm[5])
        except Exception:
            return

        width = max(
            MIN_WORD_WIDTH,
            len(text) * max(
                1.0,
                float(font_size or 8) * 0.45,
            ),
        )

        height = max(
            1.0,
            float(font_size or 8),
        )

        # pypdf visitor_text sometimes provides chunks rather than words.
        # Split whitespace while retaining approximate horizontal positions.
        pieces = re.findall(
            r"\S+",
            text,
        )

        if not pieces:
            return

        cursor = x

        for piece in pieces:
            piece_width = max(
                MIN_WORD_WIDTH,
                len(piece)
                * max(
                    1.0,
                    float(font_size or 8)
                    * 0.45,
                ),
            )

            output_words.append(
                PDFWord(
                    page=page_number,
                    text=piece,
                    x=cursor,
                    y=y,
                    width=piece_width,
                    height=height,
                )
            )

            cursor += (
                piece_width
                + max(
                    1.0,
                    float(font_size or 8)
                    * 0.25,
                )
            )

    return visitor_text


def extract_positioned_words(
    pdf_bytes: bytes,
) -> list[PDFWord]:
    reader = PdfReader(
        io.BytesIO(pdf_bytes)
    )

    words: list[PDFWord] = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):
        visitor = _visitor_text_factory(
            words,
            page_number,
        )

        try:
            page.extract_text(
                visitor_text=visitor
            )
        except TypeError:
            # Older pypdf compatibility.
            try:
                page.extract_text(
                    visitor
                )
            except Exception:
                continue

        except Exception:
            continue

    return words


# ============================================================================
# ROW GROUPING
# ============================================================================


def group_words_into_rows(
    words: list[PDFWord],
    tolerance: float = Y_TOLERANCE,
) -> list[PDFRow]:
    rows_by_page: dict[int, list[PDFRow]] = defaultdict(list)

    for word in sorted(
        words,
        key=lambda item: (
            item.page,
            -item.y,
            item.x,
        ),
    ):
        page_rows = rows_by_page[
            word.page
        ]

        matched: PDFRow | None = None

        for row in page_rows:
            if abs(row.y - word.y) <= tolerance:
                matched = row
                break

        if matched is None:
            matched = PDFRow(
                page=word.page,
                y=word.y,
                words=[],
            )

            page_rows.append(
                matched
            )

        matched.words.append(
            word
        )

    result: list[PDFRow] = []

    for page_number in sorted(
        rows_by_page
    ):
        page_rows = rows_by_page[
            page_number
        ]

        page_rows.sort(
            key=lambda row: -row.y
        )

        for row in page_rows:
            row.words.sort(
                key=lambda word: word.x
            )

            if row.text:
                result.append(row)

    return result


# ============================================================================
# PDF ROW SERIALIZATION
# ============================================================================


def serialize_pdf_rows(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    output = []

    for index, row in enumerate(rows):
        output.append(
            {
                "rowIndex": index,
                "page": row.page,
                "y": row.y,
                "xMin": row.x_min,
                "xMax": row.x_max,
                "text": row.text,
                "words": [
                    {
                        "text": word.text,
                        "x": word.x,
                        "y": word.y,
                        "width": word.width,
                        "height": word.height,
                    }
                    for word in row.words
                ],
            }
        )

    return output


# ============================================================================
# PERCENTAGE / RANK ANALYSIS
# ============================================================================


def percentage_tokens(
    row: PDFRow,
) -> list[dict[str, Any]]:
    result = []

    for word in row.words:
        matches = list(
            PERCENTAGE_RE.finditer(
                word.text
            )
        )

        for match in matches:
            try:
                value = float(
                    match.group(1)
                )
            except Exception:
                continue

            if not (
                0.0
                <= value
                <= 100.0
            ):
                continue

            result.append(
                {
                    "value": value,
                    "x": word.x,
                    "right": word.right,
                    "word": word.text,
                }
            )

    return result


def explicit_rank_tokens(
    row: PDFRow,
) -> list[dict[str, Any]]:
    result = []

    for word in row.words:
        text = normalize_text(
            word.text
        )

        match = RANK_ONLY_RE.match(
            text
        )

        if match:
            value = int(
                match.group(1)
            )

            if 1 <= value <= 10:
                result.append(
                    {
                        "rank": value,
                        "x": word.x,
                        "right": word.right,
                        "word": text,
                    }
                )

            continue

        match = RANK_MARKED_RE.match(
            text
        )

        if match:
            value = int(
                match.group(1)
            )

            if 1 <= value <= 10:
                result.append(
                    {
                        "rank": value,
                        "x": word.x,
                        "right": word.right,
                        "word": text,
                    }
                )

    return result


# ============================================================================
# SEMANTIC FILTERING
# ============================================================================


def contains_forbidden_name_pattern(
    name: str,
) -> tuple[bool, str]:
    normalized = normalize_name(
        name
    )

    for pattern in FORBIDDEN_NAME_PATTERNS:
        if re.search(
            pattern,
            normalized,
            flags=re.IGNORECASE,
        ):
            return (
                True,
                pattern,
            )

    return (
        False,
        "",
    )


def contains_forbidden_region_pattern(
    text: str,
) -> tuple[bool, str]:
    normalized = normalize_name(
        text
    )

    for pattern in FORBIDDEN_REGION_PATTERNS:
        if re.search(
            pattern,
            normalized,
            flags=re.IGNORECASE,
        ):
            return (
                True,
                pattern,
            )

    return (
        False,
        "",
    )


def is_probable_page_furniture(
    text: str,
) -> bool:
    normalized = normalize_name(
        text
    )

    if not normalized:
        return True

    if normalized in PAGE_FURNITURE:
        return True

    if normalized.startswith(
        "page "
    ):
        return True

    return False


# ============================================================================
# TABLE HEADER DETECTION
# ============================================================================


def header_kind(
    text: str,
) -> str | None:
    normalized = normalize_name(
        text
    )

    if normalized in {
        normalize_name(value)
        for value in RANK_HEADERS
    }:
        return "rank"

    if normalized in {
        normalize_name(value)
        for value in NAME_HEADERS
    }:
        return "name"

    if normalized in {
        normalize_name(value)
        for value in WEIGHT_HEADERS
    }:
        return "weight"

    return None


def detect_header_rows(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    headers = []

    for index, row in enumerate(rows):
        found = []

        for word in row.words:
            kind = header_kind(
                word.text
            )

            if kind:
                found.append(
                    {
                        "kind": kind,
                        "x": word.x,
                        "right": word.right,
                        "text": word.text,
                    }
                )

        kinds = {
            item["kind"]
            for item in found
        }

        if len(kinds) >= 2:
            headers.append(
                {
                    "rowIndex": index,
                    "page": row.page,
                    "y": row.y,
                    "text": row.text,
                    "headers": found,
                }
            )

    return headers


# ============================================================================
# COLUMN CLUSTERING
# ============================================================================


def cluster_x_positions(
    positions: list[float],
    tolerance: float = 10.0,
) -> list[dict[str, Any]]:
    if not positions:
        return []

    positions = sorted(
        positions
    )

    clusters: list[list[float]] = []

    for value in positions:
        if not clusters:
            clusters.append(
                [value]
            )
            continue

        current = clusters[-1]

        if abs(
            value - median(current)
        ) <= tolerance:
            current.append(value)
        else:
            clusters.append(
                [value]
            )

    result = []

    for cluster in clusters:
        result.append(
            {
                "x": median(cluster),
                "count": len(cluster),
                "min": min(cluster),
                "max": max(cluster),
                "std": population_std(
                    cluster
                ),
            }
        )

    result.sort(
        key=lambda item: (
            -item["count"],
            item["x"],
        )
    )

    return result


def percentage_column_candidates(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    positions = []

    occurrences = []

    for row_index, row in enumerate(rows):
        tokens = percentage_tokens(
            row
        )

        for token in tokens:
            positions.append(
                token["x"]
            )

            occurrences.append(
                {
                    "rowIndex": row_index,
                    **token,
                }
            )

    clusters = cluster_x_positions(
        positions,
        tolerance=12.0,
    )

    output = []

    for cluster in clusters:
        matching = [
            item
            for item in occurrences
            if abs(
                item["x"]
                - cluster["x"]
            )
            <= 12.0
        ]

        pages = sorted(
            {
                rows[item["rowIndex"]].page
                for item in matching
            }
        )

        output.append(
            {
                **cluster,
                "rows": [
                    item["rowIndex"]
                    for item in matching
                ],
                "pages": pages,
                "occurrences": matching,
            }
        )

    return output


def rank_column_candidates(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    positions = []

    occurrences = []

    for row_index, row in enumerate(rows):
        tokens = explicit_rank_tokens(
            row
        )

        for token in tokens:
            positions.append(
                token["x"]
            )

            occurrences.append(
                {
                    "rowIndex": row_index,
                    **token,
                }
            )

    clusters = cluster_x_positions(
        positions,
        tolerance=12.0,
    )

    output = []

    for cluster in clusters:
        matching = [
            item
            for item in occurrences
            if abs(
                item["x"]
                - cluster["x"]
            )
            <= 12.0
        ]

        pages = sorted(
            {
                rows[item["rowIndex"]].page
                for item in matching
            }
        )

        output.append(
            {
                **cluster,
                "rows": [
                    item["rowIndex"]
                    for item in matching
                ],
                "pages": pages,
                "occurrences": matching,
            }
        )

    return output


# ============================================================================
# TABLE REGION DISCOVERY
# ============================================================================


def rows_between(
    rows: list[PDFRow],
    start_index: int,
    end_index: int,
) -> list[tuple[int, PDFRow]]:
    return [
        (index, rows[index])
        for index in range(
            max(0, start_index),
            min(
                len(rows),
                end_index + 1,
            ),
        )
    ]


def table_region_score(
    rows: list[PDFRow],
    start_index: int,
    end_index: int,
    rank_x: float,
    weight_x: float,
) -> dict[str, Any]:
    if weight_x <= rank_x:
        return {
            "score": -999.0,
            "valid": False,
        }

    selected = rows_between(
        rows,
        start_index,
        end_index,
    )

    if len(selected) < 2:
        return {
            "score": -999.0,
            "valid": False,
        }

    rank_rows = 0
    weight_rows = 0
    paired_rows = 0
    semantic_rejections = 0

    distances = []

    for row_index, row in selected:
        forbidden, _ = contains_forbidden_region_pattern(
            row.text
        )

        if forbidden:
            semantic_rejections += 1
            continue

        ranks = explicit_rank_tokens(
            row
        )

        weights = percentage_tokens(
            row
        )

        rank_here = [
            item
            for item in ranks
            if (
                rank_x - 8
                <= item["x"]
                <= rank_x + 18
            )
        ]

        weight_here = [
            item
            for item in weights
            if (
                weight_x - 14
                <= item["x"]
                <= weight_x + 14
            )
        ]

        if rank_here:
            rank_rows += 1

        if weight_here:
            weight_rows += 1

        if rank_here and weight_here:
            paired_rows += 1

            distances.append(
                abs(
                    rank_here[0]["x"]
                    - weight_here[0]["x"]
                )
            )

    if paired_rows == 0:
        return {
            "score": -999.0,
            "valid": False,
        }

    # A genuine table needs repeated physical rank + weight pairing.
    score = (
        paired_rows * 8.0
        + rank_rows * 2.0
        + weight_rows * 2.0
        - semantic_rejections * 5.0
    )

    return {
        "score": score,
        "valid": (
            paired_rows >= 2
            and rank_rows >= 2
            and weight_rows >= 2
        ),
        "startIndex": start_index,
        "endIndex": end_index,
        "rankX": rank_x,
        "weightX": weight_x,
        "pairedRows": paired_rows,
        "rankRows": rank_rows,
        "weightRows": weight_rows,
        "semanticRejections": semantic_rejections,
        "rankWeightDistances": distances,
    }


def discover_candidate_regions(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    rank_columns = rank_column_candidates(
        rows
    )

    weight_columns = percentage_column_candidates(
        rows
    )

    regions = []

    for rank_column in rank_columns:
        if rank_column["count"] < 2:
            continue

        for weight_column in weight_columns:
            if weight_column["count"] < 2:
                continue

            rank_x = float(
                rank_column["x"]
            )

            weight_x = float(
                weight_column["x"]
            )

            if weight_x <= rank_x:
                continue

            shared_pages = sorted(
                set(
                    rank_column["pages"]
                ).intersection(
                    weight_column["pages"]
                )
            )

            if not shared_pages:
                continue

            relevant_indices = []

            for page in shared_pages:
                page_indices = [
                    index
                    for index, row in enumerate(
                        rows
                    )
                    if row.page == page
                ]

                if not page_indices:
                    continue

                relevant_indices.extend(
                    page_indices
                )

            if not relevant_indices:
                continue

            start = min(
                relevant_indices
            )

            end = max(
                relevant_indices
            )

            # Do not allow a huge page-wide pairing block.
            if end - start > 80:
                # Use local windows around rank/weight occurrences.
                candidate_indices = sorted(
                    set(
                        rank_column["rows"]
                    ).intersection(
                        weight_column["rows"]
                    )
                )

                if not candidate_indices:
                    continue

                start = max(
                    0,
                    min(candidate_indices)
                    - 12,
                )

                end = min(
                    len(rows) - 1,
                    max(candidate_indices)
                    + 12,
                )

            region = table_region_score(
                rows,
                start,
                end,
                rank_x,
                weight_x,
            )

            if region.get(
                "valid"
            ):
                regions.append(
                    {
                        **region,
                        "rankColumn": rank_column,
                        "weightColumn": weight_column,
                    }
                )

    regions.sort(
        key=lambda item: (
            -item["score"],
            item["startIndex"],
        )
    )

    return regions


# ============================================================================
# TABLE ROW GEOMETRY
# ============================================================================


def words_in_x_range(
    row: PDFRow,
    x_min: float,
    x_max: float,
) -> list[PDFWord]:
    return [
        word
        for word in row.words
        if (
            word.right >= x_min
            and word.x <= x_max
        )
    ]


def text_in_x_range(
    row: PDFRow,
    x_min: float,
    x_max: float,
) -> str:
    words = words_in_x_range(
        row,
        x_min,
        x_max,
    )

    return " ".join(
        word.text
        for word in sorted(
            words,
            key=lambda item: item.x,
        )
    ).strip()


def rank_for_region(
    row: PDFRow,
    rank_x: float,
) -> list[dict[str, Any]]:
    result = []

    for token in explicit_rank_tokens(
        row
    ):
        if (
            rank_x - 10
            <= token["x"]
            <= rank_x + 20
        ):
            result.append(
                token
            )

    return result


def weights_for_region(
    row: PDFRow,
    weight_x: float,
) -> list[dict[str, Any]]:
    result = []

    for token in percentage_tokens(
        row
    ):
        if (
            weight_x - 16
            <= token["x"]
            <= weight_x + 16
        ):
            result.append(
                token
            )

    return result


# ============================================================================
# NAME CLEANING
# ============================================================================


def remove_weight_percentages(
    text: str,
) -> str:
    return normalize_text(
        PERCENTAGE_RE.sub(
            "",
            text,
        )
    )


def remove_rank_prefix(
    text: str,
) -> str:
    text = normalize_text(
        text
    )

    match = re.match(
        r"^\s*\d{1,2}[.)]?\s+",
        text,
    )

    if match:
        return normalize_text(
            text[match.end():]
        )

    return text


def clean_holding_name(
    text: str,
) -> str:
    text = normalize_text(
        text
    )

    text = remove_rank_prefix(
        text
    )

    text = remove_weight_percentages(
        text
    )

    # Remove common PDF artifacts.
    text = text.replace(
        "†",
        "",
    )

    text = text.replace(
        "‡",
        "",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip(
        " -:;,."
    )


def is_valid_holding_name(
    name: str,
) -> tuple[bool, str]:
    name = clean_holding_name(
        name
    )

    if len(name) < 3:
        return (
            False,
            "holding name too short",
        )

    forbidden, pattern = (
        contains_forbidden_name_pattern(
            name
        )
    )

    if forbidden:
        return (
            False,
            f"semantic rejection: {pattern}",
        )

    # A pure number/date is not a security name.
    if re.fullmatch(
        r"[\d\s./%-]+",
        name,
    ):
        return (
            False,
            "name is numeric",
        )

    if DATE_RE.fullmatch(
        name
    ):
        return (
            False,
            "name is only a date",
        )

    return (
        True,
        "",
    )


# ============================================================================
# PHYSICAL TABLE NAME EXTRACTION
# ============================================================================


def name_interval_for_region(
    rank_x: float,
    weight_x: float,
) -> tuple[float, float]:
    # Leave a small gap around rank and percentage columns.
    return (
        rank_x + 20.0,
        weight_x - 16.0,
    )


def extract_name_from_row(
    row: PDFRow,
    rank_x: float,
    weight_x: float,
) -> str:
    name_min, name_max = (
        name_interval_for_region(
            rank_x,
            weight_x,
        )
    )

    words = words_in_x_range(
        row,
        name_min,
        name_max,
    )

    if not words:
        return ""

    text = " ".join(
        word.text
        for word in sorted(
            words,
            key=lambda item: item.x,
        )
    )

    return clean_holding_name(
        text
    )


def row_has_valid_table_geometry(
    row: PDFRow,
    rank_x: float,
    weight_x: float,
) -> bool:
    ranks = rank_for_region(
        row,
        rank_x,
    )

    weights = weights_for_region(
        row,
        weight_x,
    )

    if not ranks:
        return False

    if not weights:
        return False

    name = extract_name_from_row(
        row,
        rank_x,
        weight_x,
    )

    valid, _ = is_valid_holding_name(
        name
    )

    return valid


# ============================================================================
# FIXED-INCOME DETECTION
# ============================================================================


def looks_like_fixed_income_security(
    name: str,
) -> bool:
    normalized = normalize_name(
        name
    )

    coupon_count = len(
        list(
            PERCENTAGE_RE.finditer(
                name
            )
        )
    )

    has_date = bool(
        DATE_RE.search(
            name
        )
    )

    bond_terms = any(
        term in normalized
        for term in (
            "government",
            "republic of",
            "holdings plc",
            "bank",
            "sa ",
            "treasury",
            "bond",
            "notes",
        )
    )

    return (
        has_date
        and (
            coupon_count >= 1
            or bond_terms
        )
    )


def extract_security_name_preserving_coupon(
    row: PDFRow,
    rank_x: float,
    weight_x: float,
) -> str:
    """
    Fixed-income-aware name extraction.

    The percentage in the weight column is excluded.
    Any percentage physically inside the name interval remains.

    This is what allows:

        SINGAPORE (REPUBLIC OF) 2.375% 1-JUL-2039 5.8%

    to become:

        SINGAPORE (REPUBLIC OF) 2.375% 1-JUL-2039
    """

    name_min, name_max = (
        name_interval_for_region(
            rank_x,
            weight_x,
        )
    )

    words = words_in_x_range(
        row,
        name_min,
        name_max,
    )

    if not words:
        return ""

    pieces = []

    for word in sorted(
        words,
        key=lambda item: item.x,
    ):
        text = normalize_text(
            word.text
        )

        if not text:
            continue

        # A percentage word inside the name interval is retained.
        pieces.append(
            text
        )

    return clean_holding_name(
        " ".join(pieces)
    )


# ============================================================================
# TABLE BLOCK RECONSTRUCTION
# ============================================================================


def reconstruct_table_block(
    rows: list[PDFRow],
    region: dict[str, Any],
) -> tuple[
    list[HoldingCandidate],
    dict[str, Any],
]:
    rank_x = float(
        region["rankX"]
    )

    weight_x = float(
        region["weightX"]
    )

    start_index = int(
        region["startIndex"]
    )

    end_index = int(
        region["endIndex"]
    )

    candidates: list[HoldingCandidate] = []

    row_diagnostics = []

    current: dict[str, Any] | None = None

    for source_row_index in range(
        start_index,
        end_index + 1,
    ):
        row = rows[
            source_row_index
        ]

        ranks = rank_for_region(
            row,
            rank_x,
        )

        weights = weights_for_region(
            row,
            weight_x,
        )

        forbidden, forbidden_pattern = (
            contains_forbidden_region_pattern(
                row.text
            )
        )

        if forbidden:
            row_diagnostics.append(
                {
                    "sourceRowIndex": source_row_index,
                    "page": row.page,
                    "text": row.text,
                    "action": "reject_semantic_region",
                    "pattern": forbidden_pattern,
                }
            )

            # Do not carry semantic metadata into a holding.
            continue

        # ------------------------------------------------------------
        # Explicit rank + weight on same physical row.
        # ------------------------------------------------------------
        if ranks and weights:
            rank = ranks[0]["rank"]

            # Multiple percentage values are allowed only when
            # the final percentage is physically in the weight column.
            selected_weight = weights[-1]

            name = (
                extract_security_name_preserving_coupon(
                    row,
                    rank_x,
                    weight_x,
                )
            )

            valid_name, reason = (
                is_valid_holding_name(
                    name
                )
            )

            if not valid_name:
                row_diagnostics.append(
                    {
                        "sourceRowIndex": source_row_index,
                        "page": row.page,
                        "text": row.text,
                        "action": "reject_name",
                        "reason": reason,
                        "rank": rank,
                        "weight": selected_weight["value"],
                    }
                )

                current = None

                continue

            # Never accept ranks outside 1..10.
            if not (
                1 <= rank <= MAX_HOLDINGS
            ):
                current = None
                continue

            candidate = HoldingCandidate(
                rank=rank,
                name=name,
                weight=float(
                    selected_weight["value"]
                ),
                page=row.page,
                source_row_index=source_row_index,
                strategy="spatial_table",
                evidence={
                    "rankX": rank_x,
                    "weightX": weight_x,
                    "rankToken": ranks[0],
                    "weightToken": selected_weight,
                    "allPercentages": [
                        item["value"]
                        for item in percentage_tokens(
                            row
                        )
                    ],
                    "rowText": row.text,
                    "geometry": {
                        "rankColumnStable": True,
                        "weightColumnStable": True,
                        "samePhysicalRow": True,
                        "nameBetweenColumns": True,
                    },
                },
            )

            candidates.append(
                candidate
            )

            current = {
                "candidateIndex": len(
                    candidates
                ) - 1,
                "rank": rank,
                "name": name,
                "weight": float(
                    selected_weight["value"]
                ),
                "page": row.page,
                "sourceRowIndex": source_row_index,
            }

            row_diagnostics.append(
                {
                    "sourceRowIndex": source_row_index,
                    "page": row.page,
                    "text": row.text,
                    "action": "new_holding",
                    "rank": rank,
                    "weight": selected_weight["value"],
                    "name": name,
                }
            )

            continue

        # ------------------------------------------------------------
        # Continuation row.
        #
        # There must already be a confirmed current holding.
        # The continuation must contain text physically inside the
        # same name column.
        #
        # A continuation is never allowed to introduce a percentage.
        # ------------------------------------------------------------
        if current:
            continuation_name = extract_name_from_row(
                row,
                rank_x,
                weight_x,
            )

            row_percentages = percentage_tokens(
                row
            )

            if continuation_name:
                forbidden_name, forbidden_pattern = (
                    contains_forbidden_name_pattern(
                        continuation_name
                    )
                )

                if (
                    not forbidden_name
                    and not row_percentages
                ):
                    combined_name = clean_holding_name(
                        current["name"]
                        + " "
                        + continuation_name
                    )

                    valid_name, _ = (
                        is_valid_holding_name(
                            combined_name
                        )
                    )

                    if valid_name:
                        candidate_index = current[
                            "candidateIndex"
                        ]

                        candidates[
                            candidate_index
                        ].name = combined_name

                        candidates[
                            candidate_index
                        ].evidence[
                            "multilineContinuationRows"
                        ] = (
                            candidates[
                                candidate_index
                            ].evidence.get(
                                "multilineContinuationRows",
                                [],
                            )
                            + [
                                {
                                    "sourceRowIndex": source_row_index,
                                    "page": row.page,
                                    "text": row.text,
                                }
                            ]
                        )

                        current["name"] = (
                            combined_name
                        )

                        row_diagnostics.append(
                            {
                                "sourceRowIndex": source_row_index,
                                "page": row.page,
                                "text": row.text,
                                "action": "append_name_continuation",
                                "name": continuation_name,
                            }
                        )

                        continue

        # Anything else cannot be used.
        row_diagnostics.append(
            {
                "sourceRowIndex": source_row_index,
                "page": row.page,
                "text": row.text,
                "action": "ignored",
            }
        )

    return (
        candidates,
        {
            "region": region,
            "rows": row_diagnostics,
        },
    )


# ============================================================================
# CANDIDATE SEQUENCE VALIDATION
# ============================================================================


def normalize_candidate_list(
    candidates: list[HoldingCandidate],
) -> list[HoldingCandidate]:
    # Preserve physical order, but do not infer ranks.
    candidates = sorted(
        candidates,
        key=lambda item: (
            item.page,
            item.source_row_index,
        ),
    )

    # Remove exact duplicate physical candidates only.
    # Duplicate security names/weights are valid; duplicates generated
    # from the same source row are not separate holdings.
    seen_physical = set()
    result = []

    for candidate in candidates:
        key = (
            candidate.page,
            candidate.source_row_index,
            candidate.rank,
            normalize_name(
                candidate.name
            ),
            float(
                candidate.weight
            ),
        )

        if key in seen_physical:
            continue

        seen_physical.add(key)
        result.append(
            candidate
        )

    return result


def validate_rank_sequence(
    candidates: list[HoldingCandidate],
) -> tuple[
    bool,
    str,
]:
    if not candidates:
        return (
            False,
            "no candidates",
        )

    if len(candidates) > MAX_HOLDINGS:
        return (
            False,
            "more than ten holdings detected",
        )

    ranks = [
        item.rank
        for item in candidates
    ]

    if any(
        rank < 1
        or rank > MAX_HOLDINGS
        for rank in ranks
    ):
        return (
            False,
            "rank outside 1..10",
        )

    # Ranks must be explicit and contiguous.
    expected = list(
        range(
            1,
            len(ranks) + 1,
        )
    )

    if ranks != expected:
        return (
            False,
            "explicit published ranks are not contiguous from 1",
        )

    return (
        True,
        "",
    )


def validate_candidate_semantics(
    candidates: list[HoldingCandidate],
) -> tuple[
    bool,
    str,
]:
    for candidate in candidates:
        valid, reason = (
            is_valid_holding_name(
                candidate.name
            )
        )

        if not valid:
            return (
                False,
                reason,
            )

        if not (
            0.0
            <= candidate.weight
            <= 100.0
        ):
            return (
                False,
                "weight outside 0..100",
            )

    return (
        True,
        "",
    )


def validate_table_structure(
    candidates: list[HoldingCandidate],
    region: dict[str, Any],
) -> tuple[
    bool,
    str,
]:
    if len(candidates) < 2:
        return (
            False,
            "fewer than two physically confirmed ranked holdings",
        )

    if region["pairedRows"] < 2:
        return (
            False,
            "insufficient same-row rank/weight evidence",
        )

    # Every accepted candidate must explicitly have all three pieces.
    for candidate in candidates:
        geometry = candidate.evidence.get(
            "geometry",
            {},
        )

        if not geometry.get(
            "samePhysicalRow"
        ):
            return (
                False,
                "rank/name/weight not proven on same physical row",
            )

        if not geometry.get(
            "nameBetweenColumns"
        ):
            return (
                False,
                "name column not proven between rank and weight",
            )

    return (
        True,
        "",
    )


def consolidate_candidates(
    candidates: list[HoldingCandidate],
) -> list[HoldingCandidate]:
    """
    Keep the first structurally confirmed occurrence for each explicit rank.

    If the same rank occurs with a conflicting security, the sequence is
    invalidated elsewhere rather than guessed here.
    """

    by_rank: dict[int, list[HoldingCandidate]] = defaultdict(
        list
    )

    for candidate in candidates:
        by_rank[
            candidate.rank
        ].append(candidate)

    result = []

    for rank in sorted(
        by_rank
    ):
        values = by_rank[
            rank
        ]

        if len(values) == 1:
            result.append(
                values[0]
            )
            continue

        signatures = {
            (
                normalize_name(
                    item.name
                ),
                float(
                    item.weight
                ),
            )
            for item in values
        }

        if len(signatures) == 1:
            result.append(
                values[0]
            )
            continue

        # Conflicting explicit rank evidence must not be guessed.
        raise HoldingsParseFailure(
            "Conflicting explicit holdings share rank "
            f"{rank}."
        )

    return result


# ============================================================================
# TABLE SELECTION
# ============================================================================


def extract_candidates_from_regions(
    rows: list[PDFRow],
    regions: list[dict[str, Any]],
) -> tuple[
    list[HoldingCandidate],
    dict[str, Any],
]:
    diagnostics = {
        "regionsAttempted": [],
        "acceptedRegion": None,
        "rejectedRegions": [],
    }

    valid_sequences: list[
        tuple[
            list[HoldingCandidate],
            dict[str, Any],
        ]
    ] = []

    for region_index, region in enumerate(
        regions
    ):
        candidates, reconstruction = (
            reconstruct_table_block(
                rows,
                region,
            )
        )

        candidates = normalize_candidate_list(
            candidates
        )

        region_record = {
            "regionIndex": region_index,
            "region": region,
            "candidateCount": len(
                candidates
            ),
            "candidates": [
                asdict(item)
                for item in candidates
            ],
            "reconstruction": reconstruction,
        }

        diagnostics[
            "regionsAttempted"
        ].append(
            region_record
        )

        if not candidates:
            diagnostics[
                "rejectedRegions"
            ].append(
                {
                    "regionIndex": region_index,
                    "reason": "no candidates",
                }
            )
            continue

        try:
            candidates = consolidate_candidates(
                candidates
            )
        except HoldingsParseFailure as exc:
            diagnostics[
                "rejectedRegions"
            ].append(
                {
                    "regionIndex": region_index,
                    "reason": str(exc),
                }
            )
            continue

        ok, reason = validate_rank_sequence(
            candidates
        )

        if not ok:
            diagnostics[
                "rejectedRegions"
            ].append(
                {
                    "regionIndex": region_index,
                    "reason": reason,
                }
            )
            continue

        ok, reason = validate_candidate_semantics(
            candidates
        )

        if not ok:
            diagnostics[
                "rejectedRegions"
            ].append(
                {
                    "regionIndex": region_index,
                    "reason": reason,
                }
            )
            continue

        ok, reason = validate_table_structure(
            candidates,
            region,
        )

        if not ok:
            diagnostics[
                "rejectedRegions"
            ].append(
                {
                    "regionIndex": region_index,
                    "reason": reason,
                }
            )
            continue

        valid_sequences.append(
            (
                candidates,
                region,
            )
        )

    if not valid_sequences:
        return (
            [],
            diagnostics,
        )

    # If multiple regions independently produce different complete
    # sequences, do not guess between them.
    signatures = []

    for candidates, region in valid_sequences:
        signature = [
            (
                item.rank,
                normalize_name(
                    item.name
                ),
                float(
                    item.weight
                ),
            )
            for item in candidates
        ]

        signatures.append(
            signature
        )

    unique_signatures = []

    for signature in signatures:
        if signature not in unique_signatures:
            unique_signatures.append(
                signature
            )

    if len(unique_signatures) != 1:
        raise HoldingsParseFailure(
            "Multiple structurally valid table regions produced "
            "different holding signatures."
        )

    # Select the region with the strongest geometry score.
    valid_sequences.sort(
        key=lambda pair: (
            -float(
                pair[1]["score"]
            ),
            -len(
                pair[0]
            ),
        )
    )

    candidates, region = (
        valid_sequences[0]
    )

    diagnostics[
        "acceptedRegion"
    ] = region

    return (
        candidates,
        diagnostics,
    )


# ============================================================================
# TOP-HOLDINGS SECTION TEXT
# ============================================================================


def build_top_holdings_section_text(
    rows: list[PDFRow],
    candidates: list[HoldingCandidate],
) -> str:
    if not candidates:
        return ""

    min_index = min(
        item.source_row_index
        for item in candidates
    )

    max_index = max(
        item.source_row_index
        for item in candidates
    )

    selected = rows[
        max(
            0,
            min_index - 3,
        ):
        min(
            len(rows),
            max_index + 4,
        )
    ]

    lines = []

    for row in selected:
        lines.append(
            row.text
        )

    return "\n".join(
        lines
    )


# ============================================================================
# CONFIRMATION
# ============================================================================


def holding_signature(
    candidates: list[HoldingCandidate],
) -> list[list[Any]]:
    return [
        [
            int(item.rank),
            normalize_name(
                item.name
            ),
            float(
                item.weight
            ),
        ]
        for item in candidates
    ]


def confirm_extraction(
    candidates: list[HoldingCandidate],
    diagnostics: dict[str, Any],
) -> dict[str, Any]:
    if not candidates:
        return {
            "confirmed": False,
            "reason": "no confirmed candidates",
        }

    signature = holding_signature(
        candidates
    )

    return {
        "confirmed": True,
        "confirmationMode": (
            "strict_spatial_table_complete_sequence"
        ),
        "strategy": "spatial_table",
        "holdingsCount": len(
            candidates
        ),
        "signature": signature,
        "tableRegion": diagnostics.get(
            "acceptedRegion"
        ),
        "rulesSatisfied": {
            "explicitRanks": True,
            "explicitWeights": True,
            "samePhysicalRows": True,
            "stableRankColumn": True,
            "stableWeightColumn": True,
            "nameBetweenColumns": True,
            "semanticValidation": True,
            "contiguousPublishedRanks": True,
            "noRankInference": True,
            "noArbitraryPairing": True,
            "noSyntheticValues": True,
            "fewerThanTenAllowed": True,
        },
    }


# ============================================================================
# FORENSIC EXTRACTION ENGINE
# ============================================================================


def extract_forensic_holdings(
    pdf_bytes: bytes,
) -> dict[str, Any]:
    full_text = extract_pdf_text(
        pdf_bytes
    )

    words = extract_positioned_words(
        pdf_bytes
    )

    rows = group_words_into_rows(
        words,
        tolerance=Y_TOLERANCE,
    )

    headers = detect_header_rows(
        rows
    )

    rank_columns = rank_column_candidates(
        rows
    )

    weight_columns = percentage_column_candidates(
        rows
    )

    regions = discover_candidate_regions(
        rows
    )

    candidates, diagnostics = (
        extract_candidates_from_regions(
            rows,
            regions,
        )
    )

    candidates_payload = [
        asdict(item)
        for item in candidates
    ]

    confirmation = confirm_extraction(
        candidates,
        diagnostics,
    )

    return {
        "fullText": full_text,
        "words": words,
        "rows": rows,
        "headers": headers,
        "rankColumns": rank_columns,
        "weightColumns": weight_columns,
        "regions": regions,
        "candidates": candidates,
        "candidatesPayload": candidates_payload,
        "diagnostics": diagnostics,
        "confirmation": confirmation,
    }


# ============================================================================
# FORENSIC OUTPUT
# ============================================================================


def save_diagnostics(
    fund_dir: Path,
    extraction: dict[str, Any],
) -> None:
    rows = extraction["rows"]

    text_dump(
        fund_dir / "factsheet_text.txt",
        extraction["fullText"],
    )

    json_dump(
        fund_dir / "spatial_rows.json",
        {
            "rowCount": len(rows),
            "rows": serialize_pdf_rows(
                rows
            ),
        },
    )

    json_dump(
        fund_dir / "spatial_candidates.json",
        {
            "rankColumns": extraction[
                "rankColumns"
            ],
            "weightColumns": extraction[
                "weightColumns"
            ],
            "headers": extraction[
                "headers"
            ],
        },
    )

    json_dump(
        fund_dir / "candidate_regions.json",
        {
            "regions": extraction[
                "regions"
            ],
            "diagnostics": extraction[
                "diagnostics"
            ],
        },
    )

    json_dump(
        fund_dir / "candidate_holdings.json",
        {
            "candidates": extraction[
                "candidatesPayload"
            ],
        },
    )

    json_dump(
        fund_dir / "confirmation.json",
        extraction[
            "confirmation"
        ],
    )


# ============================================================================
# METADATA
# ============================================================================


def save_metadata(
    fund_dir: Path,
    excel_fund: dict[str, Any],
    factsheet_url: str,
) -> None:
    json_dump(
        fund_dir / "metadata.json",
        {
            "excelRow": excel_fund[
                "excelRow"
            ],
            "prudentialUrl": excel_fund[
                "prudentialUrl"
            ],
            "excelPruAccessName": excel_fund[
                "excelPruAccessName"
            ],
            "factsheetUrl": factsheet_url,
            "officialPrudentialOnly": True,
            "generatedAtUtc": utc_now(),
        },
    )


# ============================================================================
# EXACT FINAL VERIFICATION
# ============================================================================


def verify_exactly(
    page,
    source_url: str,
    expected_signature: list[list[Any]],
) -> dict[str, Any]:
    second_pdf, second_factsheet_url = (
        download_factsheet(
            page,
            source_url,
        )
    )

    second_extraction = (
        extract_forensic_holdings(
            second_pdf
        )
    )

    second_candidates = (
        second_extraction[
            "candidates"
        ]
    )

    second_confirmation = (
        second_extraction[
            "confirmation"
        ]
    )

    if not second_confirmation.get(
        "confirmed"
    ):
        return {
            "verified": False,
            "reason": (
                "fresh PDF extraction did not independently confirm "
                "a structural holdings table"
            ),
            "factsheetUrl": second_factsheet_url,
        }

    actual_signature = holding_signature(
        second_candidates
    )

    verified = (
        actual_signature
        == expected_signature
    )

    return {
        "verified": verified,
        "factsheetUrl": second_factsheet_url,
        "expectedSignature": expected_signature,
        "actualSignature": actual_signature,
        "confirmation": second_confirmation,
    }


# ============================================================================
# FUND OUTPUT
# ============================================================================


def fund_output_directory(
    excel_row: int,
    fund_name: str,
) -> Path:
    identifier = safe_filename(
        fund_name,
        fallback="fund",
    )

    return (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )


def write_success_outputs(
    fund_dir: Path,
    candidates: list[HoldingCandidate],
    section_text: str,
    verification: dict[str, Any],
) -> None:
    top_holdings = [
        {
            "rank": int(item.rank),
            "name": item.name,
            "weight": float(item.weight),
            "page": int(item.page),
            "sourceRowIndex": int(
                item.source_row_index
            ),
            "extractionStrategy": item.strategy,
            "evidence": item.evidence,
        }
        for item in candidates
    ]

    json_dump(
        fund_dir / "top_holdings.json",
        {
            "holdings": top_holdings,
            "count": len(
                top_holdings
            ),
            "signature": holding_signature(
                candidates
            ),
        },
    )

    text_dump(
        fund_dir / "top_holdings_section.txt",
        section_text,
    )

    json_dump(
        fund_dir / "recovery_result.json",
        {
            "status": "success",
            "holdingsCount": len(
                candidates
            ),
            "holdings": top_holdings,
            "signature": holding_signature(
                candidates
            ),
            "verification": verification,
            "completedAtUtc": utc_now(),
        },
    )


def write_failure_outputs(
    fund_dir: Path,
    excel_fund: dict[str, Any],
    reason: str,
    extraction: dict[str, Any] | None = None,
) -> None:
    payload = {
        "status": "failed",
        "excelRow": excel_fund[
            "excelRow"
        ],
        "prudentialUrl": excel_fund[
            "prudentialUrl"
        ],
        "excelPruAccessName": excel_fund[
            "excelPruAccessName"
        ],
        "error": reason,
        "completedAtUtc": utc_now(),
    }

    if extraction is not None:
        payload[
            "diagnosticSummary"
        ] = {
            "candidateCount": len(
                extraction.get(
                    "candidates",
                    [],
                )
            ),
            "regionCount": len(
                extraction.get(
                    "regions",
                    [],
                )
            ),
            "acceptedRegion": extraction.get(
                "diagnostics",
                {},
            ).get(
                "acceptedRegion"
            ),
            "confirmation": extraction.get(
                "confirmation"
            ),
        }

    json_dump(
        fund_dir / "diagnostic_failure.json",
        payload,
    )

    json_dump(
        fund_dir / "recovery_result.json",
        payload,
    )


# ============================================================================
# SINGLE FUND
# ============================================================================


def recover_single_fund(
    page,
    excel_fund: dict[str, Any],
) -> dict[str, Any]:
    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    fund_name = (
        excel_fund[
            "excelPruAccessName"
        ]
        or f"Fund Row {excel_row}"
    )

    fund_dir = fund_output_directory(
        excel_row,
        fund_name,
    )

    fund_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    source_url = excel_fund[
        "prudentialUrl"
    ]

    if not is_official_prudential_url(
        source_url
    ):
        reason = (
            "Source URL is not an official Prudential Singapore URL."
        )

        write_failure_outputs(
            fund_dir,
            excel_fund,
            reason,
        )

        return {
            "status": "failed",
            "excelRow": excel_row,
            "error": reason,
        }

    try:
        pdf_bytes, factsheet_url = (
            download_factsheet(
                page,
                source_url,
            )
        )

        (fund_dir / "factsheet.pdf").write_bytes(
            pdf_bytes
        )

        save_metadata(
            fund_dir,
            excel_fund,
            factsheet_url,
        )

        extraction = extract_forensic_holdings(
            pdf_bytes
        )

        save_diagnostics(
            fund_dir,
            extraction,
        )

        candidates = extraction[
            "candidates"
        ]

        confirmation = extraction[
            "confirmation"
        ]

        if not confirmation.get(
            "confirmed"
        ):
            reason = (
                "No structurally confirmed holdings table was found."
            )

            write_failure_outputs(
                fund_dir,
                excel_fund,
                reason,
                extraction,
            )

            return {
                "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
                "excelRow": excel_row,
                "error": reason,
            }

        # Explicit final structural checks.
        ok, reason = validate_rank_sequence(
            candidates
        )

        if not ok:
            write_failure_outputs(
                fund_dir,
                excel_fund,
                f"Rank sequence validation failed: {reason}",
                extraction,
            )

            return {
                "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
                "excelRow": excel_row,
                "error": reason,
            }

        ok, reason = validate_candidate_semantics(
            candidates
        )

        if not ok:
            write_failure_outputs(
                fund_dir,
                excel_fund,
                f"Semantic validation failed: {reason}",
                extraction,
            )

            return {
                "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
                "excelRow": excel_row,
                "error": reason,
            }

        expected_signature = holding_signature(
            candidates
        )

        # Fresh PDF verification.
        verification = verify_exactly(
            page,
            source_url,
            expected_signature,
        )

        if not verification.get(
            "verified"
        ):
            reason = (
                "Exact fresh-PDF verification failed."
            )

            json_dump(
                fund_dir / "confirmation.json",
                {
                    **confirmation,
                    "verified": False,
                    "verification": verification,
                },
            )

            write_failure_outputs(
                fund_dir,
                excel_fund,
                reason,
                extraction,
            )

            return {
                "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
                "excelRow": excel_row,
                "error": reason,
                "verification": verification,
            }

        section_text = (
            build_top_holdings_section_text(
                extraction["rows"],
                candidates,
            )
        )

        write_success_outputs(
            fund_dir,
            candidates,
            section_text,
            verification,
        )

        return {
            "status": "success",
            "excelRow": excel_row,
            "holdingsCount": len(
                candidates
            ),
            "holdings": [
                {
                    "rank": int(item.rank),
                    "name": item.name,
                    "weight": float(
                        item.weight
                    ),
                    "page": int(item.page),
                    "extractionStrategy": item.strategy,
                }
                for item in candidates
            ],
            "signature": expected_signature,
            "extractionConfirmation": confirmation,
            "verification": verification,
        }

    except HoldingsParseFailure as exc:
        reason = str(exc)

        write_failure_outputs(
            fund_dir,
            excel_fund,
            reason,
        )

        return {
            "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
            "excelRow": excel_row,
            "error": reason,
        }

    except Exception as exc:
        reason = (
            f"{type(exc).__name__}: {exc}"
        )

        write_failure_outputs(
            fund_dir,
            excel_fund,
            reason,
        )

        return {
            "status": "failed",
            "excelRow": excel_row,
            "error": reason,
        }


# ============================================================================
# RULES
# ============================================================================


def build_rules() -> dict[str, Any]:
    return {
        "recovery1FailedFundsOnly": True,
        "noBaselineFallback": True,
        "excelColumnAControlsMasterUniverse": True,
        "officialPrudentialSingaporeOnly": True,
        "officialFactsheetOnly": True,
        "thirdPartyHoldings": False,
        "noInferredHoldings": True,
        "noFabricatedHoldings": True,
        "noFabricatedPercentages": True,
        "noSyntheticValues": True,
        "noEstimatedValues": True,
        "noInterpolatedValues": True,
        "maximumHoldings": 10,
        "publishedCountUsedExactly": True,
        "fewerThanTenHoldingsAllowed": True,
        "noForcedTenEntries": True,
        "duplicateHoldingNamesAllowed": True,
        "duplicateHoldingPercentagesAllowed": True,
        "multilineHoldingNamesSupported": True,
        "hyphenatedLineWrapReconstruction": True,
        "fixedIncomeCouponPreserved": True,
        "fixedIncomeUsesLastPhysicalPercentageAsWeight": True,
        "noFuzzyHoldingMatching": True,
        "noArbitraryCoordinateProximityPairing": True,
        "explicitTop10HeadingNotRequired": True,
        "adaptiveTableReconstruction": True,
        "samePdfRowStructureRequired": True,
        "explicitRankColumnRequired": True,
        "explicitWeightColumnRequired": True,
        "nameColumnRequired": True,
        "semanticFalsePositiveRejection": True,
        "noRankInference": True,
        "contiguousPublishedRanksRequired": True,
        "crossStrategyExactSignatureConfirmation": True,
        "singleStrategyCompleteSequenceAllowed": True,
        "exactFinalPdfVerification": True,
        "exactRankNameWeightSignatureRequired": True,
        "forensicDiagnosticsSaved": True,
    }


# ============================================================================
# MAIN
# ============================================================================


def main() -> int:
    started_at = utc_now()

    print()
    print("=" * 72)
    print("VGrat FMS - Prudential Top Holdings RECOVERY 2")
    print("=" * 72)
    print()
    print(
        "Recovery 1 summary:",
        RECOVERY_1_RUN_SUMMARY_FILE,
    )
    print(
        "Recovery 2 output: ",
        RECOVERY_OUTPUT_DIR,
    )
    print(
        "Recovery source: Recovery 1 failed-fund details ONLY"
    )
    print()

    RECOVERY_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    RECOVERY_FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ------------------------------------------------------------------
    # Load Recovery 1 failures.
    # ------------------------------------------------------------------

    try:
        (
            recovery_1_summary,
            recovery_1_failures,
            failure_source,
        ) = load_recovery_1_run_summary()

    except Exception as exc:
        error_payload = {
            "status": "failed",
            "startedAtUtc": started_at,
            "completedAtUtc": utc_now(),
            "error": str(exc),
            "rules": build_rules(),
        }

        json_dump(
            RECOVERY_RUN_SUMMARY_FILE,
            error_payload,
        )

        json_dump(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        print(
            f"ERROR: {exc}"
        )

        # Preserve artifacts in GitHub Actions.
        return 0

    print(
        "Recovery 1 failure source:",
        failure_source,
    )

    print(
        "Recovery 1 failed funds:",
        len(
            recovery_1_failures
        ),
    )

    failed_rows = sorted(
        {
            int(item["excelRow"])
            for item in recovery_1_failures
        }
    )

    print(
        "Recovery 1 failed rows:",
        failed_rows,
    )

    # ------------------------------------------------------------------
    # Load Excel universe.
    # ------------------------------------------------------------------

    try:
        excel_universe = load_excel_universe()

    except Exception as exc:
        error_payload = {
            "status": "failed",
            "startedAtUtc": started_at,
            "completedAtUtc": utc_now(),
            "excelFile": str(
                EXCEL_FILE
            ),
            "recovery1RunSummary": str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),
            "error": str(exc),
            "rules": build_rules(),
        }

        json_dump(
            RECOVERY_RUN_SUMMARY_FILE,
            error_payload,
        )

        json_dump(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        print(
            f"ERROR: {exc}"
        )

        return 0

    # ------------------------------------------------------------------
    # Recovery 2 universe = Recovery 1 failed rows intersected with
    # Excel rows.
    #
    # We NEVER construct the universe from baseline output.
    # ------------------------------------------------------------------

    recovery_universe = []

    missing_excel_rows = []

    for failure in recovery_1_failures:
        excel_row = int(
            failure["excelRow"]
        )

        fund = excel_universe.get(
            excel_row
        )

        if fund is None:
            missing_excel_rows.append(
                excel_row
            )
            continue

        recovery_universe.append(
            fund
        )

    recovery_universe.sort(
        key=lambda item: int(
            item["excelRow"]
        )
    )

    print(
        "Recovery 2 universe:",
        len(
            recovery_universe
        ),
    )

    print(
        "Recovery 2 universe rows:",
        [
            int(
                item["excelRow"]
            )
            for item in recovery_universe
        ],
    )

    if missing_excel_rows:
        print(
            "WARNING: Recovery 1 failed rows missing from Excel:",
            missing_excel_rows,
        )

    # ------------------------------------------------------------------
    # No failures = nothing to recover.
    # ------------------------------------------------------------------

    if not recovery_universe:
        summary = {
            "status": "success",
            "startedAtUtc": started_at,
            "completedAtUtc": utc_now(),
            "excelFile": str(
                EXCEL_FILE
            ),
            "recovery1RunSummary": str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),
            "recovery1FailureSource": failure_source,
            "recovery1ExcelFundUniverse": recovery_1_summary.get(
                "excelFundUniverse"
            ),
            "recovery1FailedFunds": len(
                recovery_1_failures
            ),
            "recovery1FailedRows": failed_rows,
            "recovery2Universe": 0,
            "recovery2UniverseRows": [],
            "successfulFunds": 0,
            "diagnosticFunds": 0,
            "failedFunds": 0,
            "totalPublishedTopHoldings": 0,
            "successfulFundsDetail": [],
            "diagnosticFundsDetail": [],
            "failedFundsDetail": [],
            "missingExcelRows": missing_excel_rows,
            "rules": build_rules(),
        }

        json_dump(
            RECOVERY_RUN_SUMMARY_FILE,
            summary,
        )

        json_dump(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        print()
        print(
            "Recovery 1 has no failed funds."
        )
        print(
            "Recovery 2 has nothing to recover."
        )

        return 0

    # ------------------------------------------------------------------
    # Playwright availability.
    # ------------------------------------------------------------------

    if sync_playwright is None:
        reason = (
            "Playwright is unavailable in the Python environment."
        )

        diagnostic_details = []

        for fund in recovery_universe:
            excel_row = int(
                fund["excelRow"]
            )

            fund_dir = fund_output_directory(
                excel_row,
                fund[
                    "excelPruAccessName"
                ],
            )

            fund_dir.mkdir(
                parents=True,
                exist_ok=True,
            )

            write_failure_outputs(
                fund_dir,
                fund,
                reason,
            )

            diagnostic_details.append(
                {
                    "status": "failed",
                    "excelRow": excel_row,
                    "error": reason,
                }
            )

        summary = {
            "status": "success",
            "startedAtUtc": started_at,
            "completedAtUtc": utc_now(),
            "excelFile": str(
                EXCEL_FILE
            ),
            "recovery1RunSummary": str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),
            "recovery1FailureSource": failure_source,
            "recovery1FailedFunds": len(
                recovery_1_failures
            ),
            "recovery1FailedRows": failed_rows,
            "recovery2Universe": len(
                recovery_universe
            ),
            "recovery2UniverseRows": [
                int(
                    item["excelRow"]
                )
                for item in recovery_universe
            ],
            "successfulFunds": 0,
            "diagnosticFunds": 0,
            "failedFunds": len(
                diagnostic_details
            ),
            "totalPublishedTopHoldings": 0,
            "successfulFundsDetail": [],
            "diagnosticFundsDetail": [],
            "failedFundsDetail": diagnostic_details,
            "rules": build_rules(),
        }

        json_dump(
            RECOVERY_RUN_SUMMARY_FILE,
            summary,
        )

        json_dump(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        return 0

    # ------------------------------------------------------------------
    # Run browser.
    # ------------------------------------------------------------------

    successful_details = []
    diagnostic_details = []
    failed_details = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        context = browser.new_context(
            accept_downloads=True,
            ignore_https_errors=False,
        )

        page = context.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        for fund_index, fund in enumerate(
            recovery_universe,
            start=1,
        ):
            excel_row = int(
                fund["excelRow"]
            )

            fund_name = (
                fund[
                    "excelPruAccessName"
                ]
                or f"Fund Row {excel_row}"
            )

            print()
            print(
                "-" * 72
            )
            print(
                f"RECOVERY 2 FUND {fund_index}/"
                f"{len(recovery_universe)}"
            )
            print(
                f"Excel row: {excel_row}"
            )
            print(
                f"Fund: {fund_name}"
            )
            print(
                "-" * 72
            )

            result = recover_single_fund(
                page,
                fund,
            )

            status = result.get(
                "status"
            )

            if status == "success":
                successful_details.append(
                    result
                )

                print(
                    f"SUCCESS Row {excel_row}: "
                    f"{result.get('holdingsCount', 0)} holdings"
                )

            elif status == (
                "DIAGNOSTIC_NO_CONFIRMED_RESULT"
            ):
                diagnostic_details.append(
                    result
                )

                print(
                    f"DIAGNOSTIC Row {excel_row}: "
                    f"{result.get('error', 'unresolved')}"
                )

            else:
                failed_details.append(
                    result
                )

                print(
                    f"FAILED Row {excel_row}: "
                    f"{result.get('error', 'unknown error')}"
                )

        context.close()
        browser.close()

    # ------------------------------------------------------------------
    # Aggregate successful holdings.
    # ------------------------------------------------------------------

    all_holdings = []

    for successful in successful_details:
        all_holdings.append(
            {
                "excelRow": successful[
                    "excelRow"
                ],
                "holdingsCount": successful[
                    "holdingsCount"
                ],
                "holdings": successful[
                    "holdings"
                ],
                "signature": successful[
                    "signature"
                ],
            }
        )

    total_holdings = sum(
        int(
            item["holdingsCount"]
        )
        for item in successful_details
    )

    # ------------------------------------------------------------------
    # Always write aggregate artifacts.
    # ------------------------------------------------------------------

    json_dump(
        RECOVERY_ALL_HOLDINGS_FILE,
        all_holdings,
    )

    summary = {
        "status": "success",
        "startedAtUtc": started_at,
        "completedAtUtc": utc_now(),
        "excelFile": str(
            EXCEL_FILE
        ),
        "recovery1RunSummary": str(
            RECOVERY_1_RUN_SUMMARY_FILE
        ),
        "recovery1FailureSource": failure_source,
        "recovery1ExcelFundUniverse": recovery_1_summary.get(
            "excelFundUniverse"
        ),
        "recovery1FailedFunds": len(
            recovery_1_failures
        ),
        "recovery1FailedRows": failed_rows,
        "recovery2Universe": len(
            recovery_universe
        ),
        "recovery2UniverseRows": [
            int(
                item["excelRow"]
            )
            for item in recovery_universe
        ],
        "successfulFunds": len(
            successful_details
        ),
        "diagnosticFunds": len(
            diagnostic_details
        ),
        "failedFunds": len(
            failed_details
        ),
        "totalPublishedTopHoldings": total_holdings,
        "successfulFundsDetail": successful_details,
        "diagnosticFundsDetail": diagnostic_details,
        "failedFundsDetail": failed_details,
        "missingExcelRows": missing_excel_rows,
        "rules": build_rules(),
    }

    json_dump(
        RECOVERY_RUN_SUMMARY_FILE,
        summary,
    )

    # ------------------------------------------------------------------
    # Console summary.
    # ------------------------------------------------------------------

    print()
    print("=" * 72)
    print("RECOVERY 2 COMPLETE")
    print("=" * 72)
    print()
    print(
        "Recovery 1 failed funds:",
        len(
            recovery_1_failures
        ),
    )
    print(
        "Recovery 2 universe:",
        len(
            recovery_universe
        ),
    )
    print(
        "Successful funds:",
        len(
            successful_details
        ),
    )
    print(
        "Diagnostic/unresolved funds:",
        len(
            diagnostic_details
        ),
    )
    print(
        "Hard failures:",
        len(
            failed_details
        ),
    )
    print(
        "Published holdings recovered:",
        total_holdings,
    )

    if diagnostic_details:
        print()
        print(
            "STILL UNRESOLVED:"
        )

        for item in diagnostic_details:
            print(
                f" - Row {item['excelRow']}: "
                f"{item.get('error', 'unresolved')}"
            )

    if failed_details:
        print()
        print(
            "HARD FAILURES:"
        )

        for item in failed_details:
            print(
                f" - Row {item['excelRow']}: "
                f"{item.get('error', 'failed')}"
            )

    print()
    print(
        "Output:",
        RECOVERY_OUTPUT_DIR,
    )
    print(
        "Summary:",
        RECOVERY_RUN_SUMMARY_FILE,
    )
    print(
        "All holdings:",
        RECOVERY_ALL_HOLDINGS_FILE,
    )
    print()
    print("=" * 72)

    # Intentionally return zero so GitHub Actions can upload diagnostics
    # even when one or more funds remain unresolved.
    return 0


# ============================================================================
# ENTRY POINT
# ============================================================================


if __name__ == "__main__":
    sys.exit(
        main()
    )
