#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings RECOVERY 2
================================================

RECOVERY 2 MASTER RULE
======================

This script operates ONLY on funds that FAILED in Recovery 1.

Recovery chain:

    Funds Links.xlsm
          |
          v
    test_prudential_holdings.py
          |
          v
    update_prudential_holdings.py
          |
          v
    output_holdings_recovery/run_summary.json
          |
          v
    THIS SCRIPT
    test_prudential_holdings_recovery2.py

IMPORTANT
=========

This script MUST NOT:

- read baseline failures directly
- add new funds to the universe
- modify baseline output
- modify Recovery 1 output
- use third-party holdings data
- infer holdings
- fabricate percentages
- estimate percentages
- interpolate percentages
- fuzzy-match security names
- arbitrarily pair unrelated coordinates
- force exactly 10 holdings

This script is deliberately forensic.

It first downloads the official Prudential Singapore factsheet and
reconstructs the PDF's table structure using multiple independent views:

1. normal PDF text
2. positioned PDF words
3. positioned PDF rows
4. rank candidates
5. percentage candidates
6. table-column geometry
7. candidate holding blocks

The explicit "Top 10 Holdings" heading is NOT required.

A successful extraction must still satisfy strict validation:

- sequential ranks beginning at 1
- maximum 10 holdings
- every holding has a real published percentage
- no missing names
- no missing weights
- weights are 0-100
- no arbitrary coordinate pairing
- duplicate names are allowed
- fewer than 10 holdings are allowed
- exact signature survives a fresh official-PDF verification

For fixed-income lines containing multiple percentages, the LAST
percentage on the same logical holding row/block is treated as the
published portfolio weight, consistent with the Recovery 1 rules.
"""

from __future__ import annotations

import io
import json
import re
import statistics
import sys
import time
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import sync_playwright


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

Y_TOLERANCE = 3.5

MIN_WORD_WIDTH = 0.5

PERCENTAGE_RE = re.compile(
    r"(?<![\d.])(\d+(?:\.\d+)?)\s*%"
)

RANK_RE = re.compile(
    r"^\s*(\d{1,2})(?:[.)]|\s+)"
)

RANK_ONLY_RE = re.compile(
    r"^\s*(\d{1,2})\s*$"
)

NOISE_RE = re.compile(
    r"""
    ^
    (?:
        top\s+(?:10|ten)\s+holdings
        |holdings
        |holding
        |portfolio
        |portfolio\s+holdings
        |investment\s+holdings
        |investments
        |investment
        |security
        |securities
        |name
        |weight
        |%
        |source
        |important\s+information
        |disclaimer
        |past\s+performance
        |asset\s+allocation
        |portfolio\s+characteristics
        |fund\s+information
        |fund\s+facts
        |factsheet
        |page\s+\d+
        |\d{1,2}\s+months?
        |\d{1,2}\s+years?
    )
    $
    """,
    re.IGNORECASE | re.VERBOSE,
)

KEYWORD_RE = re.compile(
    r"\b("
    r"holdings?|"
    r"portfolio|"
    r"investment|"
    r"investments|"
    r"security|"
    r"securities|"
    r"equity|"
    r"bond|"
    r"fund|"
    r"company|"
    r"corporation|"
    r"limited|"
    r"plc|"
    r"inc\.?|"
    r"ltd\.?"
    r")\b",
    re.IGNORECASE,
)


# ============================================================================
# DATA STRUCTURES
# ============================================================================


@dataclass
class PDFWord:
    page: int
    text: str
    x: float
    y: float
    width: float
    height: float
    font_size: float

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def center_x(self) -> float:
        return self.x + self.width / 2.0

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2.0


@dataclass
class PDFRow:
    page: int
    y: float
    words: list[PDFWord]

    @property
    def x_min(self) -> float:
        if not self.words:
            return 0.0

        return min(
            word.x
            for word in self.words
        )

    @property
    def x_max(self) -> float:
        if not self.words:
            return 0.0

        return max(
            word.right
            for word in self.words
        )

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
    raw_row: str
    evidence: dict[str, Any]


class HoldingsParseFailure(Exception):
    pass


# ============================================================================
# GENERAL HELPERS
# ============================================================================


def utc_now_iso() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = (
        text.replace("\u00a0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\u2010", "-")
        .replace("\u2011", "-")
        .replace("\u2012", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\ufb01", "fi")
        .replace("\ufb02", "fl")
    )

    text = re.sub(
        r"[ \t\r\f\v]+",
        " ",
        text,
    )

    text = re.sub(
        r"\n+",
        "\n",
        text,
    )

    return text.strip()


def normalize_name_for_signature(
    name: str,
) -> str:
    value = normalize_text(
        name
    ).lower()

    value = value.replace(
        "’",
        "'",
    )

    value = value.replace(
        "“",
        '"',
    )

    value = value.replace(
        "”",
        '"',
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    value = re.sub(
        r"\s*-\s*",
        "-",
        value,
    )

    return value.strip()


def safe_filename(
    value: str,
    max_length: int = 150,
) -> str:
    value = normalize_text(
        value
    )

    value = re.sub(
        r"[^A-Za-z0-9._()\- ]+",
        "_",
        value,
    )

    value = re.sub(
        r"\s+",
        "_",
        value,
    )

    value = re.sub(
        r"_+",
        "_",
        value,
    )

    value = value.strip(
        "._"
    )

    if not value:
        value = "fund"

    return value[:max_length]


def ensure_dir(
    path: Path,
) -> None:
    path.mkdir(
        parents=True,
        exist_ok=True,
    )


def write_json(
    path: Path,
    payload: Any,
) -> None:
    ensure_dir(
        path.parent
    )

    path.write_text(
        json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def write_text(
    path: Path,
    text: str,
) -> None:
    ensure_dir(
        path.parent
    )

    path.write_text(
        text,
        encoding="utf-8",
    )


def is_prudential_url(
    url: str,
) -> bool:
    try:
        host = (
            urlparse(url)
            .netloc
            .lower()
            .split(":")[0]
        )
    except Exception:
        return False

    return host in PRUDENTIAL_HOSTS


def sleep_retry() -> None:
    time.sleep(
        RETRY_DELAY_SECONDS
    )


# ============================================================================
# RECOVERY 1 FAILURE LOADING
# ============================================================================


def normalize_failed_fund_entries(
    candidate: Any,
) -> list[dict[str, Any]]:
    if not isinstance(
        candidate,
        list,
    ):
        return []

    result: list[
        dict[str, Any]
    ] = []

    for item in candidate:
        if not isinstance(
            item,
            dict,
        ):
            continue

        row = item.get(
            "excelRow"
        )

        try:
            if row is not None:
                row = int(row)
        except Exception:
            continue

        if row is None:
            continue

        copied = dict(
            item
        )

        copied[
            "excelRow"
        ] = row

        result.append(
            copied
        )

    result.sort(
        key=lambda item: int(
            item["excelRow"]
        )
    )

    return result


def load_recovery_1_failure_files(
) -> list[dict[str, Any]]:
    base = (
        Path(
            "output_holdings_recovery"
        )
        / "funds"
    )

    if not base.exists():
        return []

    recovered: list[
        dict[str, Any]
    ] = []

    for failure_file in sorted(
        base.glob(
            "*_failed/failure.json"
        ),
        key=lambda path: path.parent.name,
    ):
        try:
            payload = json.loads(
                failure_file.read_text(
                    encoding="utf-8"
                )
            )
        except Exception:
            continue

        if not isinstance(
            payload,
            dict,
        ):
            continue

        status = normalize_text(
            payload.get(
                "status"
            )
        ).lower()

        if status != "failed":
            continue

        excel_row = payload.get(
            "excelRow"
        )

        if excel_row is None:
            match = re.match(
                r"^(\d+)_failed$",
                failure_file.parent.name,
            )

            if match:
                excel_row = int(
                    match.group(1)
                )

        try:
            excel_row = int(
                excel_row
            )
        except Exception:
            continue

        item = dict(
            payload
        )

        item[
            "excelRow"
        ] = excel_row

        item[
            "_reconstructedFromRecovery1Output"
        ] = True

        item[
            "_recovery1FailureFile"
        ] = str(
            failure_file
        )

        recovered.append(
            item
        )

    recovered.sort(
        key=lambda item: int(
            item["excelRow"]
        )
    )

    return recovered


def load_recovery_1_run_summary(
) -> tuple[
    dict[str, Any],
    str,
]:
    if not RECOVERY_1_RUN_SUMMARY_FILE.exists():
        raise RuntimeError(
            "Recovery 1 run summary does not exist: "
            f"{RECOVERY_1_RUN_SUMMARY_FILE}"
        )

    payload = json.loads(
        RECOVERY_1_RUN_SUMMARY_FILE.read_text(
            encoding="utf-8"
        )
    )

    if not isinstance(
        payload,
        dict,
    ):
        raise RuntimeError(
            "Recovery 1 run summary is not a JSON object."
        )

    possible_keys = [
        "failedFundsDetail",
        "failedFundsDetails",
        "failedFunds",
        "failedFundsList",
        "failureDetails",
        "failures",
        "failed",
        "failedFundDetails",
    ]

    for key in possible_keys:
        value = payload.get(
            key
        )

        normalized = (
            normalize_failed_fund_entries(
                value
            )
        )

        if normalized:
            return (
                payload,
                f"recovery1_summary:{key}",
            )

    file_reconstructed = (
        load_recovery_1_failure_files()
    )

    if file_reconstructed:
        return (
            payload,
            "recovery1_output_failure_files",
        )

    return (
        payload,
        "none",
    )


# ============================================================================
# EXCEL
# ============================================================================


def read_excel_funds(
) -> dict[int, dict[str, Any]]:
    if not EXCEL_FILE.exists():
        raise RuntimeError(
            f"Excel file does not exist: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_links=True,
    )

    worksheet = workbook.active

    funds: dict[
        int,
        dict[str, Any],
    ] = {}

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

        url = normalize_text(
            url_value
        )

        pru_name = normalize_text(
            name_value
        )

        if not url:
            continue

        funds[
            row_number
        ] = {
            "excelRow": row_number,
            "prudentialUrl": url,
            "excelPruAccessName": pru_name,
        }

    workbook.close()

    return funds


def build_recovery_universe(
    recovery_1_failures: list[
        dict[str, Any]
    ],
) -> list[
    dict[str, Any]
]:
    excel_funds = (
        read_excel_funds()
    )

    universe: list[
        dict[str, Any]
    ] = []

    for failure in recovery_1_failures:
        row = int(
            failure[
                "excelRow"
            ]
        )

        if row not in excel_funds:
            raise RuntimeError(
                "Recovery 1 failed fund row "
                f"{row} does not exist in "
                "Funds Links.xlsm Column A."
            )

        fund = dict(
            excel_funds[
                row
            ]
        )

        fund[
            "recovery1Failure"
        ] = failure

        universe.append(
            fund
        )

    universe.sort(
        key=lambda item: int(
            item["excelRow"]
        )
    )

    return universe


# ============================================================================
# PLAYWRIGHT / FACTSHEET
# ============================================================================


def launch_browser(
    playwright,
):
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

    return (
        browser,
        context,
        page,
    )


def open_prudential_page(
    page,
    prudential_url: str,
) -> None:
    if not is_prudential_url(
        prudential_url
    ):
        raise RuntimeError(
            "Non-Prudential URL rejected: "
            f"{prudential_url}"
        )

    last_error: Exception | None = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):
        try:
            response = page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            if response is not None:
                if response.status >= 400:
                    raise RuntimeError(
                        "Prudential page returned "
                        f"HTTP {response.status}"
                    )

            page.wait_for_timeout(
                POST_PAGE_WAIT_MS
            )

            return

        except Exception as exc:
            last_error = exc

            if attempt < RETRY_COUNT:
                sleep_retry()

    raise RuntimeError(
        "Unable to open Prudential page: "
        f"{prudential_url}; "
        f"last error: {last_error}"
    )


def find_factsheet_url(
    page,
    prudential_url: str,
) -> str:
    links = page.locator(
        "a"
    ).evaluate_all(
        """
        elements => elements.map(a => ({
            href: a.href || "",
            text: a.innerText || "",
            title: a.getAttribute("title") || "",
            aria: a.getAttribute("aria-label") || ""
        }))
        """
    )

    candidates: list[
        tuple[int, str, str]
    ] = []

    for link in links:
        href = normalize_text(
            link.get(
                "href"
            )
        )

        if not href:
            continue

        absolute = urljoin(
            prudential_url,
            href,
        )

        if not is_prudential_url(
            absolute
        ):
            continue

        combined = " ".join(
            [
                normalize_text(
                    link.get(
                        "text"
                    )
                ),
                normalize_text(
                    link.get(
                        "title"
                    )
                ),
                normalize_text(
                    link.get(
                        "aria"
                    )
                ),
                absolute,
            ]
        ).lower()

        score = 0

        if ".pdf" in absolute.lower():
            score += 100

        if "factsheet" in combined:
            score += 80

        if "fund factsheet" in combined:
            score += 40

        if "fund" in combined:
            score += 20

        if "download" in combined:
            score += 10

        if score > 0:
            candidates.append(
                (
                    score,
                    absolute,
                    combined,
                )
            )

    if not candidates:
        raise RuntimeError(
            "No official Prudential factsheet link "
            "could be identified."
        )

    candidates.sort(
        key=lambda item: (
            -item[0],
            item[1],
        )
    )

    return candidates[0][1]


def download_factsheet(
    page,
    factsheet_url: str,
) -> bytes:
    if not is_prudential_url(
        factsheet_url
    ):
        raise RuntimeError(
            "Non-Prudential factsheet URL rejected."
        )

    last_error: Exception | None = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):
        try:
            response = page.request.get(
                factsheet_url,
                timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
            )

            if response.status != 200:
                raise RuntimeError(
                    "Factsheet HTTP status "
                    f"{response.status}"
                )

            body = response.body()

            if not body.startswith(
                b"%PDF"
            ):
                raise RuntimeError(
                    "Downloaded factsheet is not a PDF."
                )

            return body

        except Exception as exc:
            last_error = exc

            if attempt < RETRY_COUNT:
                sleep_retry()

    raise RuntimeError(
        "Unable to download official factsheet: "
        f"{last_error}"
    )


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================


def extract_pdf_text(
    factsheet_bytes: bytes,
) -> str:
    reader = PdfReader(
        io.BytesIO(
            factsheet_bytes
        )
    )

    pages: list[str] = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):
        try:
            text = (
                pdf_page.extract_text()
                or ""
            )
        except Exception:
            text = ""

        pages.append(
            f"===== PAGE {page_number} =====\n"
            f"{text}"
        )

    return "\n\n".join(
        pages
    )


def pdf_lines(
    text: str,
) -> list[str]:
    lines: list[str] = []

    for raw in text.splitlines():
        value = normalize_text(
            raw
        )

        if value:
            lines.append(
                value
            )

    return lines


def extract_positioned_words(
    factsheet_bytes: bytes,
) -> list[PDFWord]:
    reader = PdfReader(
        io.BytesIO(
            factsheet_bytes
        )
    )

    words: list[
        PDFWord
    ] = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        def visitor_text(
            text,
            cm,
            tm,
            font_dict,
            font_size,
        ):
            value = normalize_text(
                text
            )

            if not value:
                return

            try:
                x = float(
                    tm[4]
                )

                y = float(
                    tm[5]
                )

            except Exception:
                return

            try:
                fs = abs(
                    float(
                        font_size or 0
                    )
                )
            except Exception:
                fs = 0.0

            chunks = re.findall(
                r"\S+",
                value,
            )

            if not chunks:
                return

            if len(chunks) == 1:
                chunk = chunks[0]

                estimated_width = max(
                    MIN_WORD_WIDTH,
                    len(chunk)
                    * max(
                        fs * 0.48,
                        2.0,
                    ),
                )

                words.append(
                    PDFWord(
                        page=page_number,
                        text=chunk,
                        x=x,
                        y=y,
                        width=estimated_width,
                        height=max(
                            fs,
                            1.0,
                        ),
                        font_size=fs,
                    )
                )

                return

            estimated_char_width = max(
                fs * 0.48,
                2.0,
            )

            offset = 0.0

            for chunk in chunks:
                width = max(
                    MIN_WORD_WIDTH,
                    len(chunk)
                    * estimated_char_width,
                )

                words.append(
                    PDFWord(
                        page=page_number,
                        text=chunk,
                        x=x + offset,
                        y=y,
                        width=width,
                        height=max(
                            fs,
                            1.0,
                        ),
                        font_size=fs,
                    )
                )

                offset += (
                    width
                    + estimated_char_width
                )

        try:
            pdf_page.extract_text(
                visitor_text=visitor_text
            )
        except Exception:
            continue

    return words


def group_words_into_rows(
    words: list[PDFWord],
    y_tolerance: float = Y_TOLERANCE,
) -> list[PDFRow]:
    by_page: dict[
        int,
        list[PDFWord],
    ] = defaultdict(list)

    for word in words:
        by_page[
            word.page
        ].append(
            word
        )

    rows: list[
        PDFRow
    ] = []

    for page_number in sorted(
        by_page
    ):
        page_words = sorted(
            by_page[
                page_number
            ],
            key=lambda word: (
                -word.y,
                word.x,
            ),
        )

        page_rows: list[
            PDFRow
        ] = []

        for word in page_words:
            matched: PDFRow | None = None

            for candidate in reversed(
                page_rows
            ):
                if abs(
                    candidate.y
                    - word.y
                ) <= y_tolerance:
                    matched = candidate
                    break

                if (
                    candidate.y
                    < word.y
                    - y_tolerance
                ):
                    break

            if matched is None:
                matched = PDFRow(
                    page=page_number,
                    y=word.y,
                    words=[],
                )

                page_rows.append(
                    matched
                )

            matched.words.append(
                word
            )

        for row in page_rows:
            row.words.sort(
                key=lambda item: item.x
            )

        page_rows.sort(
            key=lambda row: -row.y
        )

        rows.extend(
            page_rows
        )

    return rows


# ============================================================================
# TEXT / ROW ANALYSIS
# ============================================================================


def extract_percentages(
    text: str,
) -> list[float]:
    result: list[
        float
    ] = []

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
            result.append(
                value
            )

    return result


def extract_rank_from_text(
    text: str,
) -> int | None:
    match = RANK_RE.match(
        text
    )

    if not match:
        return None

    try:
        rank = int(
            match.group(1)
        )
    except Exception:
        return None

    if 1 <= rank <= MAX_HOLDINGS:
        return rank

    return None


def strip_rank(
    text: str,
) -> str:
    return re.sub(
        r"^\s*\d{1,2}(?:[.)]|\s+)",
        "",
        text,
        count=1,
    ).strip()


def clean_holding_name(
    value: str,
) -> str:
    value = normalize_text(
        value
    )

    value = re.sub(
        r"^\s*\d{1,2}(?:[.)]|\s+)",
        "",
        value,
        count=1,
    )

    value = PERCENTAGE_RE.sub(
        "",
        value,
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    value = value.strip(
        " \t-–—:;|"
    )

    return value


def is_noise_name(
    name: str,
) -> bool:
    value = normalize_text(
        name
    )

    if not value:
        return True

    if NOISE_RE.match(
        value
    ):
        return True

    if len(value) < 2:
        return True

    if PERCENTAGE_RE.fullmatch(
        value
    ):
        return True

    if re.fullmatch(
        r"[\d\s.,%-]+",
        value,
    ):
        return True

    return False


def merge_hyphenated_lines(
    lines: list[str],
) -> list[str]:
    merged: list[
        str
    ] = []

    index = 0

    while index < len(lines):
        current = normalize_text(
            lines[index]
        )

        if (
            current.endswith("-")
            and index + 1 < len(lines)
        ):
            next_line = normalize_text(
                lines[index + 1]
            )

            if next_line:
                merged.append(
                    current[:-1]
                    + next_line
                )

                index += 2
                continue

        merged.append(
            current
        )

        index += 1

    return merged


def row_has_rank(
    row: PDFRow,
) -> bool:
    return (
        extract_rank_from_text(
            row.text
        )
        is not None
    )


def row_has_percentage(
    row: PDFRow,
) -> bool:
    return bool(
        extract_percentages(
            row.text
        )
    )


# ============================================================================
# EXPLICIT SECTION DISCOVERY
# ============================================================================


def find_explicit_section_rows(
    rows: list[PDFRow],
) -> list[int]:
    matches: list[
        int
    ] = []

    for index, row in enumerate(
        rows
    ):
        text = normalize_text(
            row.text
        )

        if re.search(
            r"\btop\s+(?:10|ten)\s+holdings\b",
            text,
            flags=re.IGNORECASE,
        ):
            matches.append(
                index
            )

    return matches


# ============================================================================
# TABLE REGION DETECTION
# ============================================================================


def row_keyword_score(
    row: PDFRow,
) -> int:
    return len(
        KEYWORD_RE.findall(
            row.text
        )
    )


def find_candidate_regions(
    rows: list[PDFRow],
) -> list[dict[str, Any]]:
    regions: list[
        dict[str, Any]
    ] = []

    rows_by_page: dict[
        int,
        list[
            tuple[int, PDFRow]
        ],
    ] = defaultdict(list)

    for index, row in enumerate(
        rows
    ):
        rows_by_page[
            row.page
        ].append(
            (
                index,
                row,
            )
        )

    for page_number in sorted(
        rows_by_page
    ):
        page_rows = rows_by_page[
            page_number
        ]

        if not page_rows:
            continue

        for start_pos in range(
            len(page_rows)
        ):
            end_pos = min(
                len(page_rows),
                start_pos + 30,
            )

            window = page_rows[
                start_pos:end_pos
            ]

            rank_count = sum(
                1
                for _, row in window
                if row_has_rank(row)
            )

            percentage_count = sum(
                len(
                    extract_percentages(
                        row.text
                    )
                )
                for _, row in window
            )

            keyword_count = sum(
                row_keyword_score(
                    row
                )
                for _, row in window
            )

            if (
                rank_count >= 2
                and percentage_count >= 3
            ) or (
                rank_count >= 1
                and percentage_count >= 6
                and keyword_count >= 2
            ):
                regions.append(
                    {
                        "page": page_number,
                        "startRowIndex": window[0][0],
                        "endRowIndex": window[-1][0],
                        "rankCount": rank_count,
                        "percentageCount": percentage_count,
                        "keywordCount": keyword_count,
                        "startY": window[0][1].y,
                        "endY": window[-1][1].y,
                    }
                )

    collapsed: list[
        dict[str, Any]
    ] = []

    for region in regions:
        matched = None

        for existing in collapsed:
            if (
                existing["page"]
                == region["page"]
                and not (
                    region["endRowIndex"]
                    < existing["startRowIndex"]
                    or region["startRowIndex"]
                    > existing["endRowIndex"]
                )
            ):
                matched = existing
                break

        if matched is None:
            collapsed.append(
                dict(region)
            )
        else:
            matched[
                "startRowIndex"
            ] = min(
                matched[
                    "startRowIndex"
                ],
                region[
                    "startRowIndex"
                ],
            )

            matched[
                "endRowIndex"
            ] = max(
                matched[
                    "endRowIndex"
                ],
                region[
                    "endRowIndex"
                ],
            )

            matched[
                "rankCount"
            ] = max(
                matched[
                    "rankCount"
                ],
                region[
                    "rankCount"
                ],
            )

            matched[
                "percentageCount"
            ] = max(
                matched[
                    "percentageCount"
                ],
                region[
                    "percentageCount"
                ],
            )

            matched[
                "keywordCount"
            ] = max(
                matched[
                    "keywordCount"
                ],
                region[
                    "keywordCount"
                ],
            )

    collapsed.sort(
        key=lambda item: (
            -item["rankCount"],
            -item["percentageCount"],
            -item["keywordCount"],
            item["page"],
            item["startRowIndex"],
        )
    )

    return collapsed


# ============================================================================
# COLUMN / GEOMETRY ANALYSIS
# ============================================================================


def percentage_tokens_in_row(
    row: PDFRow,
) -> list[
    tuple[PDFWord, float]
]:
    result: list[
        tuple[PDFWord, float]
    ] = []

    for word in row.words:
        matches = list(
            PERCENTAGE_RE.finditer(
                word.text
            )
        )

        if not matches:
            continue

        for match in matches:
            try:
                value = float(
                    match.group(1)
                )
            except Exception:
                continue

            if 0 <= value <= 100:
                result.append(
                    (
                        word,
                        value,
                    )
                )

    return result


def rank_token(
    row: PDFRow,
) -> tuple[
    PDFWord,
    int,
] | None:
    for word in row.words:
        rank = extract_rank_from_text(
            word.text
        )

        if rank is not None:
            return (
                word,
                rank,
            )

    # Some PDFs expose the complete row text but fragment the rank
    # into a separate one-character object. Handle an exact rank-only
    # text object without using coordinate proximity to invent a match.
    for word in row.words:
        match = RANK_ONLY_RE.match(
            word.text
        )

        if not match:
            continue

        try:
            rank = int(
                match.group(1)
            )
        except Exception:
            continue

        if 1 <= rank <= MAX_HOLDINGS:
            return (
                word,
                rank,
            )

    return None


def estimate_percentage_columns(
    rows: list[PDFRow],
) -> list[float]:
    xs: list[
        float
    ] = []

    for row in rows:
        for word, _value in (
            percentage_tokens_in_row(
                row
            )
        ):
            xs.append(
                word.center_x
            )

    if not xs:
        return []

    xs.sort()

    clusters: list[
        list[float]
    ] = []

    for x in xs:
        if not clusters:
            clusters.append(
                [x]
            )
            continue

        if abs(
            statistics.median(
                clusters[-1]
            )
            - x
        ) <= 18.0:
            clusters[-1].append(
                x
            )
        else:
            clusters.append(
                [x]
            )

    return [
        round(
            statistics.median(
                cluster
            ),
            2,
        )
        for cluster in clusters
        if cluster
    ]


def nearest_column(
    x: float,
    columns: list[float],
    tolerance: float = 22.0,
) -> float | None:
    if not columns:
        return None

    candidate = min(
        columns,
        key=lambda value: abs(
            value - x
        ),
    )

    if abs(
        candidate - x
    ) <= tolerance:
        return candidate

    return None


def rank_to_name_text(
    row: PDFRow,
    rank_word: PDFWord,
    percentage_x: float | None,
) -> str:
    pieces: list[
        str
    ] = []

    for word in row.words:
        if word.right <= (
            rank_word.right + 2
        ):
            continue

        if (
            percentage_x is not None
            and word.x >= percentage_x - 2
        ):
            continue

        pieces.append(
            word.text
        )

    return normalize_text(
        " ".join(
            pieces
        )
    )


# ============================================================================
# STRICT TABLE RECONSTRUCTION
# ============================================================================


def reconstruct_ranked_rows(
    rows: list[PDFRow],
    strategy_name: str,
) -> list[
    HoldingCandidate
]:
    if not rows:
        return []

    candidates: list[
        HoldingCandidate
    ] = []

    pending: dict[
        int,
        dict[str, Any],
    ] = {}

    for row_index, row in enumerate(
        rows
    ):
        text = normalize_text(
            row.text
        )

        if not text:
            continue

        rank_info = rank_token(
            row
        )

        percentage_items = (
            percentage_tokens_in_row(
                row
            )
        )

        if rank_info is not None:
            rank_word, rank = (
                rank_info
            )

            if not (
                1 <= rank <= MAX_HOLDINGS
            ):
                continue

            for old_rank, state in list(
                pending.items()
            ):
                if old_rank != rank:
                    if state.get(
                        "weight"
                    ) is None and not state.get(
                        "weightValues"
                    ):
                        state[
                            "unresolvedBeforeNewRank"
                        ] = True

            name_text = rank_to_name_text(
                row,
                rank_word,
                (
                    percentage_items[-1][0].center_x
                    if percentage_items
                    else None
                ),
            )

            name_text = clean_holding_name(
                name_text
            )

            state = {
                "rank": rank,
                "nameParts": [],
                "weightValues": [],
                "page": row.page,
                "rowIndex": row_index,
                "rawRows": [text],
                "rankX": rank_word.x,
                "percentageX": None,
                "unresolvedBeforeNewRank": False,
            }

            if (
                name_text
                and not is_noise_name(
                    name_text
                )
            ):
                state[
                    "nameParts"
                ].append(
                    name_text
                )

            for word, value in (
                percentage_items
            ):
                state[
                    "weightValues"
                ].append(
                    value
                )

                if state[
                    "percentageX"
                ] is None:
                    state[
                        "percentageX"
                    ] = word.center_x

            pending[
                rank
            ] = state

        else:
            if not pending:
                continue

            latest_rank = max(
                pending.keys(),
                key=lambda rank: pending[
                    rank
                ]["rowIndex"],
            )

            state = pending[
                latest_rank
            ]

            state[
                "rawRows"
            ].append(
                text
            )

            if percentage_items:
                for word, value in (
                    percentage_items
                ):
                    state[
                        "weightValues"
                    ].append(
                        value
                    )

                    if state[
                        "percentageX"
                    ] is None:
                        state[
                            "percentageX"
                        ] = word.center_x

                continue

            continuation = clean_holding_name(
                text
            )

            if (
                continuation
                and not is_noise_name(
                    continuation
                )
                and not re.search(
                    r"\b(?:source|important information|"
                    r"disclaimer|past performance|"
                    r"asset allocation|"
                    r"portfolio characteristics)\b",
                    continuation,
                    flags=re.IGNORECASE,
                )
            ):
                state[
                    "nameParts"
                ].append(
                    continuation
                )

    for rank in sorted(
        pending
    ):
        state = pending[
            rank
        ]

        if state.get(
            "unresolvedBeforeNewRank"
        ):
            continue

        name = normalize_text(
            " ".join(
                state[
                    "nameParts"
                ]
            )
        )

        if not name:
            continue

        if is_noise_name(
            name
        ):
            continue

        weights = state[
            "weightValues"
        ]

        if not weights:
            continue

        weight = float(
            weights[-1]
        )

        if not (
            0.0 <= weight <= 100.0
        ):
            continue

        candidates.append(
            HoldingCandidate(
                rank=rank,
                name=name,
                weight=weight,
                page=int(
                    state["page"]
                ),
                source_row_index=int(
                    state["rowIndex"]
                ),
                strategy=strategy_name,
                raw_row="\n".join(
                    state[
                        "rawRows"
                    ]
                ),
                evidence={
                    "percentageValues": weights,
                    "percentageColumnX": state[
                        "percentageX"
                    ],
                    "rankX": state[
                        "rankX"
                    ],
                },
            )
        )

    return candidates


def reconstruct_same_row_table(
    rows: list[PDFRow],
) -> list[
    HoldingCandidate
]:
    candidates: list[
        HoldingCandidate
    ] = []

    for row_index, row in enumerate(
        rows
    ):
        rank_info = rank_token(
            row
        )

        if rank_info is None:
            continue

        rank_word, rank = (
            rank_info
        )

        if not (
            1 <= rank <= MAX_HOLDINGS
        ):
            continue

        percentage_items = (
            percentage_tokens_in_row(
                row
            )
        )

        if not percentage_items:
            continue

        weight_word, weight = (
            percentage_items[-1]
        )

        name = rank_to_name_text(
            row,
            rank_word,
            weight_word.center_x,
        )

        name = clean_holding_name(
            name
        )

        if not name:
            continue

        if is_noise_name(
            name
        ):
            continue

        candidates.append(
            HoldingCandidate(
                rank=rank,
                name=name,
                weight=float(
                    weight
                ),
                page=row.page,
                source_row_index=row_index,
                strategy="same_row",
                raw_row=row.text,
                evidence={
                    "rankX": rank_word.x,
                    "weightX": weight_word.center_x,
                    "allPercentages": [
                        value
                        for _word, value
                        in percentage_items
                    ],
                },
            )
        )

    return candidates


def reconstruct_column_table(
    rows: list[PDFRow],
) -> list[
    HoldingCandidate
]:
    candidates: list[
        HoldingCandidate
    ] = []

    percentage_columns = (
        estimate_percentage_columns(
            rows
        )
    )

    for row_index, row in enumerate(
        rows
    ):
        rank_info = rank_token(
            row
        )

        if rank_info is None:
            continue

        rank_word, rank = (
            rank_info
        )

        if not (
            1 <= rank <= MAX_HOLDINGS
        ):
            continue

        percentage_items = (
            percentage_tokens_in_row(
                row
            )
        )

        if not percentage_items:
            continue

        selected_word, selected_weight = (
            percentage_items[-1]
        )

        column = nearest_column(
            selected_word.center_x,
            percentage_columns,
            tolerance=30.0,
        )

        if column is None:
            column = selected_word.center_x

        name_parts: list[
            str
        ] = []

        for word in row.words:
            if word.right <= (
                rank_word.right + 2
            ):
                continue

            if word.x >= (
                column - 4
            ):
                continue

            if PERCENTAGE_RE.search(
                word.text
            ):
                continue

            name_parts.append(
                word.text
            )

        name = clean_holding_name(
            " ".join(
                name_parts
            )
        )

        if not name:
            continue

        if is_noise_name(
            name
        ):
            continue

        candidates.append(
            HoldingCandidate(
                rank=rank,
                name=name,
                weight=float(
                    selected_weight
                ),
                page=row.page,
                source_row_index=row_index,
                strategy="same_row_columns",
                raw_row=row.text,
                evidence={
                    "rankX": rank_word.x,
                    "weightX": selected_word.center_x,
                    "weightColumn": column,
                    "allPercentages": [
                        value
                        for _word, value
                        in percentage_items
                    ],
                },
            )
        )

    return candidates


# ============================================================================
# MULTILINE TABLE RECONSTRUCTION
# ============================================================================


def reconstruct_multiline_table(
    rows: list[PDFRow],
) -> list[
    HoldingCandidate
]:
    results: list[
        HoldingCandidate
    ] = []

    current: dict[
        str,
        Any,
    ] | None = None

    for row_index, row in enumerate(
        rows
    ):
        text = normalize_text(
            row.text
        )

        if not text:
            continue

        rank_info = rank_token(
            row
        )

        if rank_info is not None:
            if current is not None:
                candidate = (
                    finalize_multiline_state(
                        current,
                        "multiline_ranked",
                    )
                )

                if candidate is not None:
                    results.append(
                        candidate
                    )

            rank_word, rank = (
                rank_info
            )

            current = {
                "rank": rank,
                "page": row.page,
                "startRow": row_index,
                "rankX": rank_word.x,
                "nameParts": [],
                "weights": [],
                "rawRows": [],
                "weightX": None,
            }

            current[
                "rawRows"
            ].append(
                text
            )

            percentages = (
                percentage_tokens_in_row(
                    row
                )
            )

            name = rank_to_name_text(
                row,
                rank_word,
                (
                    percentages[-1][0].center_x
                    if percentages
                    else None
                ),
            )

            name = clean_holding_name(
                name
            )

            if (
                name
                and not is_noise_name(
                    name
                )
            ):
                current[
                    "nameParts"
                ].append(
                    name
                )

            for word, value in percentages:
                current[
                    "weights"
                ].append(
                    value
                )

                current[
                    "weightX"
                ] = word.center_x

            continue

        if current is None:
            continue

        current[
            "rawRows"
        ].append(
            text
        )

        percentages = (
            percentage_tokens_in_row(
                row
            )
        )

        if percentages:
            for word, value in percentages:
                current[
                    "weights"
                ].append(
                    value
                )

                current[
                    "weightX"
                ] = word.center_x

            continue

        continuation = clean_holding_name(
            text
        )

        if (
            continuation
            and not is_noise_name(
                continuation
            )
            and not re.search(
                r"\b(?:source|important information|"
                r"disclaimer|past performance|"
                r"asset allocation|"
                r"portfolio characteristics)\b",
                continuation,
                flags=re.IGNORECASE,
            )
        ):
            current[
                "nameParts"
            ].append(
                continuation
            )

    if current is not None:
        candidate = (
            finalize_multiline_state(
                current,
                "multiline_ranked",
            )
        )

        if candidate is not None:
            results.append(
                candidate
            )

    return results


def finalize_multiline_state(
    state: dict[str, Any],
    strategy: str,
) -> HoldingCandidate | None:
    rank = int(
        state[
            "rank"
        ]
    )

    if not (
        1 <= rank <= MAX_HOLDINGS
    ):
        return None

    name = normalize_text(
        " ".join(
            state[
                "nameParts"
            ]
        )
    )

    name = clean_holding_name(
        name
    )

    if not name or is_noise_name(
        name
    ):
        return None

    weights = [
        float(value)
        for value in state[
            "weights"
        ]
        if 0 <= float(value) <= 100
    ]

    if not weights:
        return None

    return HoldingCandidate(
        rank=rank,
        name=name,
        weight=weights[-1],
        page=int(
            state[
                "page"
            ]
        ),
        source_row_index=int(
            state[
                "startRow"
            ]
        ),
        strategy=strategy,
        raw_row="\n".join(
            state[
                "rawRows"
            ]
        ),
        evidence={
            "allPercentages": weights,
            "rankX": state[
                "rankX"
            ],
            "weightX": state[
                "weightX"
            ],
        },
    )


# ============================================================================
# CANDIDATE VALIDATION
# ============================================================================


def candidate_signature(
    candidates: list[
        HoldingCandidate
    ],
) -> list[
    tuple[int, str, float]
]:
    ordered = sorted(
        candidates,
        key=lambda item: item.rank,
    )

    return [
        (
            int(
                item.rank
            ),
            normalize_name_for_signature(
                item.name
            ),
            float(
                item.weight
            ),
        )
        for item in ordered
    ]


def validate_holdings(
    candidates: list[
        HoldingCandidate
    ],
) -> None:
    if not candidates:
        raise HoldingsParseFailure(
            "No holding candidates were produced."
        )

    if len(candidates) > MAX_HOLDINGS:
        raise HoldingsParseFailure(
            "Candidate extraction produced more than "
            f"{MAX_HOLDINGS} holdings."
        )

    ordered = sorted(
        candidates,
        key=lambda item: item.rank,
    )

    expected = list(
        range(
            1,
            len(ordered) + 1,
        )
    )

    actual = [
        item.rank
        for item in ordered
    ]

    if actual != expected:
        raise HoldingsParseFailure(
            "Holding ranks are not sequential: "
            f"{actual}"
        )

    seen_ranks: set[
        int
    ] = set()

    for item in ordered:
        if item.rank in seen_ranks:
            raise HoldingsParseFailure(
                "Duplicate holding rank detected: "
                f"{item.rank}"
            )

        seen_ranks.add(
            item.rank
        )

        if not item.name:
            raise HoldingsParseFailure(
                f"Holding rank {item.rank} "
                "has no name."
            )

        if not (
            0 <= item.weight <= 100
        ):
            raise HoldingsParseFailure(
                f"Holding rank {item.rank} "
                f"has invalid weight {item.weight}."
            )


def validate_candidate_structure(
    candidates: list[
        HoldingCandidate
    ],
) -> bool:
    try:
        validate_holdings(
            candidates
        )
    except HoldingsParseFailure:
        return False

    return len(candidates) >= 2


# ============================================================================
# CANDIDATE CONSOLIDATION
# ============================================================================


def deduplicate_candidates(
    candidates: list[
        HoldingCandidate
    ],
) -> list[
    HoldingCandidate
]:
    seen: set[
        tuple[int, str, float]
    ] = set()

    result: list[
        HoldingCandidate
    ] = []

    for candidate in candidates:
        key = (
            candidate.rank,
            normalize_name_for_signature(
                candidate.name
            ),
            float(
                candidate.weight
            ),
        )

        if key in seen:
            continue

        seen.add(
            key
        )

        result.append(
            candidate
        )

    return result


def sequence_key(
    candidates: list[
        HoldingCandidate
    ],
) -> tuple[
    tuple[int, str, float],
    ...,
]:
    return tuple(
        candidate_signature(
            candidates
        )
    )


def sequence_quality(
    sequence: list[
        HoldingCandidate
    ],
) -> tuple:
    if not sequence:
        return (
            -999,
        )

    ranks = [
        item.rank
        for item in sequence
    ]

    page_count = len(
        {
            item.page
            for item in sequence
        }
    )

    row_indexes = [
        item.source_row_index
        for item in sequence
    ]

    span = (
        max(row_indexes)
        - min(row_indexes)
        if row_indexes
        else 999999
    )

    strategies = Counter(
        item.strategy
        for item in sequence
    )

    same_strategy_bonus = max(
        strategies.values()
    )

    contiguous_ranks = (
        ranks
        == list(
            range(
                1,
                len(ranks) + 1,
            )
        )
    )

    return (
        len(sequence),
        1 if contiguous_ranks else 0,
        same_strategy_bonus,
        -page_count,
        -span,
    )


def group_candidate_sequences(
    candidates: list[
        HoldingCandidate
    ],
) -> list[
    list[HoldingCandidate]
]:
    by_rank: dict[
        int,
        list[HoldingCandidate],
    ] = defaultdict(list)

    for candidate in candidates:
        by_rank[
            candidate.rank
        ].append(
            candidate
        )

    if not by_rank:
        return []

    sequences: list[
        list[HoldingCandidate]
    ] = [[]]

    for rank in sorted(
        by_rank
    ):
        options = by_rank[
            rank
        ]

        new_sequences: list[
            list[HoldingCandidate]
        ] = []

        for sequence in sequences:
            for option in options:
                candidate_sequence = (
                    sequence
                    + [option]
                )

                if len(
                    candidate_sequence
                ) <= MAX_HOLDINGS:
                    new_sequences.append(
                        candidate_sequence
                    )

        sequences = new_sequences

        if len(
            sequences
        ) > 2000:
            sequences = sorted(
                sequences,
                key=sequence_quality,
                reverse=True,
            )[:500]

    return [
        sequence
        for sequence in sequences
        if validate_candidate_structure(
            sequence
        )
    ]


def choose_confirmed_sequence(
    candidates_by_strategy: dict[
        str,
        list[HoldingCandidate],
    ],
) -> tuple[
    list[HoldingCandidate] | None,
    dict[str, Any],
]:
    valid_sequences: list[
        tuple[
            str,
            list[HoldingCandidate],
        ]
    ] = []

    for strategy, candidates in (
        candidates_by_strategy.items()
    ):
        if not candidates:
            continue

        sequences = (
            group_candidate_sequences(
                candidates
            )
        )

        for sequence in sequences:
            valid_sequences.append(
                (
                    strategy,
                    sequence,
                )
            )

    if not valid_sequences:
        return (
            None,
            {
                "confirmed": False,
                "reason": (
                    "No strategy produced a "
                    "strictly valid ranked sequence."
                ),
                "strategyCount": 0,
            },
        )

    signature_groups: dict[
        tuple,
        list[
            tuple[
                str,
                list[HoldingCandidate],
            ]
        ],
    ] = defaultdict(list)

    for strategy, sequence in valid_sequences:
        signature_groups[
            sequence_key(
                sequence
            )
        ].append(
            (
                strategy,
                sequence,
            )
        )

    confirmed_groups = [
        group
        for group in signature_groups.values()
        if len(
            {
                strategy
                for strategy, _sequence
                in group
            }
        ) >= 2
    ]

    if confirmed_groups:
        confirmed_groups.sort(
            key=lambda group: sequence_quality(
                group[0][1]
            ),
            reverse=True,
        )

        best_group = (
            confirmed_groups[0]
        )

        best = max(
            best_group,
            key=lambda item: sequence_quality(
                item[1]
            ),
        )

        return (
            best[1],
            {
                "confirmed": True,
                "confirmationMode": (
                    "exact_cross_strategy_agreement"
                ),
                "strategies": [
                    strategy
                    for strategy, _sequence
                    in best_group
                ],
                "signature": [
                    list(item)
                    for item in sequence_key(
                        best[1]
                    )
                ],
            },
        )

    valid_sequences.sort(
        key=lambda item: sequence_quality(
            item[1]
        ),
        reverse=True,
    )

    best_strategy, best_sequence = (
        valid_sequences[0]
    )

    # A single strategy is allowed to confirm fewer than five holdings
    # when it has independently reconstructed a complete sequential
    # table. The exact PDF verification below provides the second layer
    # of protection.
    #
    # We still require:
    # - at least two holdings
    # - sequential ranks
    # - every holding has name + published weight
    # - no duplicate rank
    # - one coherent reconstruction strategy
    #
    # No fuzzy matching or coordinate proximity is used.
    if len(
        best_sequence
    ) >= 2:
        return (
            best_sequence,
            {
                "confirmed": True,
                "confirmationMode": (
                    "strict_single_strategy_complete_sequence"
                ),
                "strategy": best_strategy,
                "signature": [
                    list(item)
                    for item in sequence_key(
                        best_sequence
                    )
                ],
            },
        )

    return (
        None,
        {
            "confirmed": False,
            "reason": (
                "A candidate sequence existed but "
                "did not contain at least two "
                "sequential published holdings."
            ),
            "bestStrategy": best_strategy,
            "bestCount": len(
                best_sequence
            ),
            "candidateSignature": [
                list(item)
                for item in sequence_key(
                    best_sequence
                )
            ],
        },
    )


# ============================================================================
# FORENSIC DIAGNOSTICS
# ============================================================================


def build_row_diagnostics(
    rows: list[PDFRow],
) -> list[
    dict[str, Any]
]:
    result: list[
        dict[str, Any]
    ] = []

    for index, row in enumerate(
        rows
    ):
        percentages = extract_percentages(
            row.text
        )

        rank = extract_rank_from_text(
            row.text
        )

        result.append(
            {
                "rowIndex": index,
                "page": row.page,
                "y": row.y,
                "xMin": row.x_min,
                "xMax": row.x_max,
                "rank": rank,
                "percentages": percentages,
                "keywordScore": row_keyword_score(
                    row
                ),
                "text": row.text,
                "words": [
                    {
                        "text": word.text,
                        "x": word.x,
                        "y": word.y,
                        "width": word.width,
                        "height": word.height,
                        "fontSize": word.font_size,
                    }
                    for word in row.words
                ],
            }
        )

    return result


def build_text_candidate_diagnostics(
    lines: list[str],
) -> dict[str, Any]:
    rank_candidates = []

    percentage_candidates = []

    keyword_candidates = []

    for index, line in enumerate(
        lines
    ):
        rank = extract_rank_from_text(
            line
        )

        percentages = extract_percentages(
            line
        )

        keyword_matches = (
            KEYWORD_RE.findall(
                line
            )
        )

        if rank is not None:
            rank_candidates.append(
                {
                    "lineIndex": index,
                    "rank": rank,
                    "text": line,
                }
            )

        if percentages:
            percentage_candidates.append(
                {
                    "lineIndex": index,
                    "percentages": percentages,
                    "text": line,
                }
            )

        if keyword_matches:
            keyword_candidates.append(
                {
                    "lineIndex": index,
                    "keywords": keyword_matches,
                    "text": line,
                }
            )

    return {
        "rankCandidates": rank_candidates,
        "percentageCandidates": percentage_candidates,
        "keywordCandidates": keyword_candidates,
    }


def build_spatial_candidate_diagnostics(
    rows: list[PDFRow],
) -> dict[str, Any]:
    rank_candidates = []

    percentage_candidates = []

    for index, row in enumerate(
        rows
    ):
        rank_info = rank_token(
            row
        )

        if rank_info is not None:
            word, rank = (
                rank_info
            )

            rank_candidates.append(
                {
                    "rowIndex": index,
                    "page": row.page,
                    "rank": rank,
                    "x": word.x,
                    "y": word.y,
                    "text": row.text,
                }
            )

        percentages = (
            percentage_tokens_in_row(
                row
            )
        )

        for word, value in percentages:
            percentage_candidates.append(
                {
                    "rowIndex": index,
                    "page": row.page,
                    "value": value,
                    "x": word.center_x,
                    "y": word.center_y,
                    "text": row.text,
                }
            )

    return {
        "rankCandidates": rank_candidates,
        "percentageCandidates": percentage_candidates,
        "percentageColumns": (
            estimate_percentage_columns(
                rows
            )
        ),
    }


def serialize_candidate_list(
    candidates: list[
        HoldingCandidate
    ],
) -> list[
    dict[str, Any]
]:
    return [
        asdict(
            candidate
        )
        for candidate in candidates
    ]


def save_diagnostics(
    output_dir: Path,
    factsheet_bytes: bytes,
    full_text: str,
    rows: list[PDFRow],
    regions: list[dict[str, Any]],
    candidates_by_strategy: dict[
        str,
        list[HoldingCandidate],
    ],
    confirmation: dict[str, Any],
) -> None:
    ensure_dir(
        output_dir
    )

    write_text(
        output_dir
        / "factsheet_text.txt",
        full_text,
    )

    (
        output_dir
        / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    lines = pdf_lines(
        full_text
    )

    write_json(
        output_dir
        / "text_diagnostics.json",
        build_text_candidate_diagnostics(
            lines
        ),
    )

    write_json(
        output_dir
        / "spatial_rows.json",
        build_row_diagnostics(
            rows
        ),
    )

    write_json(
        output_dir
        / "spatial_candidates.json",
        build_spatial_candidate_diagnostics(
            rows
        ),
    )

    write_json(
        output_dir
        / "candidate_regions.json",
        regions,
    )

    serializable_candidates = {
        strategy: serialize_candidate_list(
            candidates
        )
        for strategy, candidates
        in candidates_by_strategy.items()
    }

    write_json(
        output_dir
        / "candidate_holdings.json",
        serializable_candidates,
    )

    write_json(
        output_dir
        / "confirmation.json",
        confirmation,
    )


# ============================================================================
# FORENSIC EXTRACTION ENGINE
# ============================================================================


def extract_forensic_holdings(
    factsheet_bytes: bytes,
) -> tuple[
    list[HoldingCandidate] | None,
    dict[str, Any],
    str,
    list[PDFRow],
    dict[
        str,
        list[HoldingCandidate],
    ],
]:
    full_text = extract_pdf_text(
        factsheet_bytes
    )

    lines = pdf_lines(
        full_text
    )

    lines = merge_hyphenated_lines(
        lines
    )

    positioned_words = (
        extract_positioned_words(
            factsheet_bytes
        )
    )

    rows = group_words_into_rows(
        positioned_words
    )

    explicit_headings = (
        find_explicit_section_rows(
            rows
        )
    )

    regions = find_candidate_regions(
        rows
    )

    region_rows: list[
        list[PDFRow]
    ] = []

    for region in regions[:20]:
        start = max(
            0,
            int(
                region[
                    "startRowIndex"
                ]
            ) - 3,
        )

        end = min(
            len(rows),
            int(
                region[
                    "endRowIndex"
                ]
            ) + 3,
        )

        subset = rows[
            start:end
        ]

        if subset:
            region_rows.append(
                subset
            )

    by_page: dict[
        int,
        list[PDFRow],
    ] = defaultdict(list)

    for row in rows:
        by_page[
            row.page
        ].append(
            row
        )

    for page_number in sorted(
        by_page
    ):
        page_rows = by_page[
            page_number
        ]

        if any(
            row_has_rank(row)
            for row in page_rows
        ) and any(
            row_has_percentage(row)
            for row in page_rows
        ):
            region_rows.append(
                page_rows
            )

    unique_regions: list[
        list[PDFRow]
    ] = []

    seen_region_keys: set[
        tuple
    ] = set()

    for subset in region_rows:
        key = tuple(
            (
                row.page,
                round(
                    row.y,
                    1,
                ),
                row.text,
            )
            for row in subset
        )

        if key in seen_region_keys:
            continue

        seen_region_keys.add(
            key
        )

        unique_regions.append(
            subset
        )

    candidates_by_strategy: dict[
        str,
        list[HoldingCandidate],
    ] = defaultdict(list)

    for subset in unique_regions:

        for candidate in (
            reconstruct_same_row_table(
                subset
            )
        ):
            candidates_by_strategy[
                "same_row"
            ].append(
                candidate
            )

        for candidate in (
            reconstruct_column_table(
                subset
            )
        ):
            candidates_by_strategy[
                "same_row_columns"
            ].append(
                candidate
            )

        for candidate in (
            reconstruct_multiline_table(
                subset
            )
        ):
            candidates_by_strategy[
                "multiline_ranked"
            ].append(
                candidate
            )

        for candidate in (
            reconstruct_ranked_rows(
                subset,
                "ranked_block",
            )
        ):
            candidates_by_strategy[
                "ranked_block"
            ].append(
                candidate
            )

    for strategy in list(
        candidates_by_strategy
    ):
        candidates_by_strategy[
            strategy
        ] = deduplicate_candidates(
            candidates_by_strategy[
                strategy
            ]
        )

    confirmed, confirmation = (
        choose_confirmed_sequence(
            candidates_by_strategy
        )
    )

    diagnostics = {
        "explicitSectionHeadings": len(
            explicit_headings
        ),
        "explicitSectionRowIndexes": (
            explicit_headings
        ),
        "totalPdfLines": len(
            lines
        ),
        "totalPositionedWords": len(
            positioned_words
        ),
        "totalSpatialRows": len(
            rows
        ),
        "candidateRegions": regions,
        "candidateRegionCount": len(
            regions
        ),
        "candidateSubsetCount": len(
            unique_regions
        ),
        "strategyCandidateCounts": {
            strategy: len(
                candidates
            )
            for strategy, candidates
            in candidates_by_strategy.items()
        },
        "confirmation": confirmation,
    }

    return (
        confirmed,
        diagnostics,
        full_text,
        rows,
        candidates_by_strategy,
    )


# ============================================================================
# HOLDING OUTPUT
# ============================================================================


def holdings_to_json(
    holdings: list[
        HoldingCandidate
    ],
) -> list[
    dict[str, Any]
]:
    ordered = sorted(
        holdings,
        key=lambda item: item.rank,
    )

    return [
        {
            "rank": int(
                item.rank
            ),
            "name": item.name,
            "weight": float(
                item.weight
            ),
            "page": int(
                item.page
            ),
            "extractionStrategy": item.strategy,
        }
        for item in ordered
    ]


def holdings_signature_json(
    holdings: list[
        HoldingCandidate
    ],
) -> list[
    list[Any]
]:
    return [
        [
            rank,
            name,
            weight,
        ]
        for rank, name, weight
        in candidate_signature(
            holdings
        )
    ]


# ============================================================================
# FINAL EXACT VERIFICATION
# ============================================================================


def verify_exact_signature(
    page,
    factsheet_url: str,
    expected_holdings: list[
        HoldingCandidate
    ],
) -> dict[str, Any]:
    fresh_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    (
        fresh_holdings,
        diagnostics,
        _text,
        _rows,
        _candidates,
    ) = extract_forensic_holdings(
        fresh_bytes
    )

    if fresh_holdings is None:
        raise HoldingsParseFailure(
            "Fresh official factsheet verification "
            "could not reconstruct the holdings."
        )

    expected_signature = (
        candidate_signature(
            expected_holdings
        )
    )

    actual_signature = (
        candidate_signature(
            fresh_holdings
        )
    )

    if (
        expected_signature
        != actual_signature
    ):
        raise HoldingsParseFailure(
            "Exact final PDF verification failed.\n"
            f"Expected: {expected_signature}\n"
            f"Actual:   {actual_signature}"
        )

    return {
        "verified": True,
        "expectedSignature": [
            list(item)
            for item in expected_signature
        ],
        "actualSignature": [
            list(item)
            for item in actual_signature
        ],
        "freshFactsheetBytes": len(
            fresh_bytes
        ),
        "freshDiagnostics": diagnostics,
    }


# ============================================================================
# FUND PROCESSING
# ============================================================================


def recover_single_fund(
    page,
    fund: dict[str, Any],
) -> dict[str, Any]:
    excel_row = int(
        fund[
            "excelRow"
        ]
    )

    prudential_url = fund[
        "prudentialUrl"
    ]

    excel_pru_name = fund[
        "excelPruAccessName"
    ]

    output_name = (
        f"{excel_row}_"
        f"{safe_filename(excel_pru_name)}"
    )

    fund_output_dir = (
        RECOVERY_FUNDS_OUTPUT_DIR
        / output_name
    )

    ensure_dir(
        fund_output_dir
    )

    started_at = utc_now_iso()

    open_prudential_page(
        page,
        prudential_url,
    )

    factsheet_url = find_factsheet_url(
        page,
        prudential_url,
    )

    factsheet_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    reader = PdfReader(
        io.BytesIO(
            factsheet_bytes
        )
    )

    page_count = len(
        reader.pages
    )

    (
        holdings,
        diagnostics,
        full_text,
        rows,
        candidates_by_strategy,
    ) = extract_forensic_holdings(
        factsheet_bytes
    )

    save_diagnostics(
        fund_output_dir,
        factsheet_bytes,
        full_text,
        rows,
        diagnostics.get(
            "candidateRegions",
            [],
        ),
        candidates_by_strategy,
        diagnostics.get(
            "confirmation",
            {},
        ),
    )

    metadata = {
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "excelPruAccessName": excel_pru_name,
        "factsheetUrl": factsheet_url,
        "factsheetPageCount": page_count,
        "recoveryStage": "Recovery 2",
        "startedAtUtc": started_at,
        "completedAtUtc": utc_now_iso(),
        "diagnostics": diagnostics,
    }

    write_json(
        fund_output_dir
        / "metadata.json",
        metadata,
    )

    if holdings is None:
        diagnostic_failure = {
            "status": (
                "DIAGNOSTIC_NO_CONFIRMED_RESULT"
            ),
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pru_name,
            "factsheetUrl": factsheet_url,
            "factsheetPageCount": page_count,
            "explicitSectionHeadings": diagnostics.get(
                "explicitSectionHeadings",
                0,
            ),
            "strategyCandidateCounts": diagnostics.get(
                "strategyCandidateCounts",
                {},
            ),
            "confirmation": diagnostics.get(
                "confirmation",
                {},
            ),
            "recovery1Failure": fund.get(
                "recovery1Failure"
            ),
            "diagnosticDirectory": str(
                fund_output_dir
            ),
            "generatedAtUtc": utc_now_iso(),
        }

        write_json(
            fund_output_dir
            / "diagnostic_failure.json",
            diagnostic_failure,
        )

        return diagnostic_failure

    validate_holdings(
        holdings
    )

    verification = (
        verify_exact_signature(
            page,
            factsheet_url,
            holdings,
        )
    )

    write_json(
        fund_output_dir
        / "top_holdings.json",
        {
            "status": "success",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pru_name,
            "factsheetUrl": factsheet_url,
            "holdings": holdings_to_json(
                holdings
            ),
            "signature": holdings_signature_json(
                holdings
            ),
            "verification": verification,
            "generatedAtUtc": utc_now_iso(),
        },
    )

    write_text(
        fund_output_dir
        / "top_holdings_section.txt",
        "\n".join(
            (
                f"{item.rank}. "
                f"{item.name} "
                f"{item.weight:.10g}%"
            )
            for item in sorted(
                holdings,
                key=lambda item: item.rank,
            )
        ),
    )

    success = {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "excelPruAccessName": excel_pru_name,
        "factsheetUrl": factsheet_url,
        "factsheetPageCount": page_count,
        "holdingsCount": len(
            holdings
        ),
        "holdings": holdings_to_json(
            holdings
        ),
        "signature": holdings_signature_json(
            holdings
        ),
        "extractionConfirmation": diagnostics.get(
            "confirmation",
            {},
        ),
        "verification": verification,
        "outputDirectory": str(
            fund_output_dir
        ),
        "generatedAtUtc": utc_now_iso(),
    }

    write_json(
        fund_output_dir
        / "recovery_result.json",
        success,
    )

    return success


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
        "maximumHoldings": MAX_HOLDINGS,
        "publishedCountUsedExactly": True,
        "fewerThanTenHoldingsAllowed": True,
        "noForcedTenEntries": True,
        "duplicateHoldingNamesAllowed": True,
        "duplicateHoldingPercentagesAllowed": True,
        "multilineHoldingNamesSupported": True,
        "hyphenatedLineWrapReconstruction": True,
        "fallbackUsesLastPercentageOnLogicalLine": True,
        "noFuzzyHoldingMatching": True,
        "noArbitraryCoordinateProximityPairing": True,
        "explicitTop10HeadingNotRequired": True,
        "adaptiveTableReconstruction": True,
        "samePdfRowStructureRequired": True,
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
    started_at = utc_now_iso()

    ensure_dir(
        RECOVERY_OUTPUT_DIR
    )

    ensure_dir(
        RECOVERY_FUNDS_OUTPUT_DIR
    )

    print()
    print(
        "=" * 72
    )
    print(
        "VGrat FMS - Prudential Top Holdings Recovery 2"
    )
    print(
        "=" * 72
    )
    print(
        "Recovery source:",
        RECOVERY_1_RUN_SUMMARY_FILE,
    )
    print(
        "Recovery output:",
        RECOVERY_OUTPUT_DIR,
    )
    print(
        "Rule: Recovery 1 failed funds ONLY"
    )
    print(
        "Rule: baseline output is NEVER used as a recovery source"
    )
    print(
        "=" * 72
    )
    print()

    (
        recovery_1_summary,
        failure_source,
    ) = load_recovery_1_run_summary()

    if failure_source.startswith(
        "recovery1_summary:"
    ):
        key = failure_source.split(
            ":",
            1,
        )[1]

        recovery_1_failures = (
            normalize_failed_fund_entries(
                recovery_1_summary.get(
                    key
                )
            )
        )

    elif (
        failure_source
        == "recovery1_output_failure_files"
    ):
        recovery_1_failures = (
            load_recovery_1_failure_files()
        )

    else:
        recovery_1_failures = []

    recovery_1_failures = sorted(
        recovery_1_failures,
        key=lambda item: int(
            item[
                "excelRow"
            ]
        ),
    )

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

    print(
        "Recovery 1 failed rows:",
        [
            int(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_1_failures
        ],
    )

    if not recovery_1_failures:
        summary = {
            "status": "success",
            "startedAtUtc": started_at,
            "completedAtUtc": utc_now_iso(),
            "excelFile": str(
                EXCEL_FILE
            ),
            "recovery1RunSummary": str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),
            "recovery1FailureSource": failure_source,
            "recovery1FailedFunds": 0,
            "recovery1FailedRows": [],
            "recovery2Universe": 0,
            "recovery2UniverseRows": [],
            "successfulFunds": 0,
            "diagnosticFunds": 0,
            "failedFunds": 0,
            "totalPublishedTopHoldings": 0,
            "successfulFundsDetail": [],
            "diagnosticFundsDetail": [],
            "failedFundsDetail": [],
            "rules": build_rules(),
        }

        write_json(
            RECOVERY_RUN_SUMMARY_FILE,
            summary,
        )

        write_json(
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

    recovery_universe = (
        build_recovery_universe(
            recovery_1_failures
        )
    )

    print(
        "Recovery 2 universe:",
        len(
            recovery_universe
        ),
    )

    print(
        "Recovery 2 rows:",
        [
            int(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_universe
        ],
    )

    successful: list[
        dict[str, Any]
    ] = []

    diagnostic: list[
        dict[str, Any]
    ] = []

    failed: list[
        dict[str, Any]
    ] = []

    all_holdings: list[
        dict[str, Any]
    ] = []

    with sync_playwright() as playwright:
        (
            browser,
            context,
            page,
        ) = launch_browser(
            playwright
        )

        try:
            for fund in recovery_universe:
                row = int(
                    fund[
                        "excelRow"
                    ]
                )

                print()
                print(
                    "-" * 72
                )
                print(
                    f"RECOVERY 2 FUND - EXCEL ROW {row}"
                )
                print(
                    "PruAccess name:",
                    fund[
                        "excelPruAccessName"
                    ],
                )
                print(
                    "URL:",
                    fund[
                        "prudentialUrl"
                    ],
                )
                print(
                    "-" * 72
                )

                try:
                    result = recover_single_fund(
                        page,
                        fund,
                    )

                    status = normalize_text(
                        result.get(
                            "status"
                        )
                    ).lower()

                    if status == "success":
                        successful.append(
                            result
                        )

                        for holding in result.get(
                            "holdings",
                            [],
                        ):
                            all_holdings.append(
                                {
                                    "excelRow": row,
                                    "prudentialUrl": fund[
                                        "prudentialUrl"
                                    ],
                                    "excelPruAccessName": fund[
                                        "excelPruAccessName"
                                    ],
                                    **holding,
                                }
                            )

                        print(
                            "STATUS: SUCCESS"
                        )

                        print(
                            "Holdings:",
                            result.get(
                                "holdingsCount"
                            ),
                        )

                    elif status.startswith(
                        "diagnostic_"
                    ):
                        diagnostic.append(
                            result
                        )

                        print(
                            "STATUS: "
                            "DIAGNOSTIC_NO_CONFIRMED_RESULT"
                        )

                        print(
                            "No strictly confirmed "
                            "holding sequence."
                        )

                    else:
                        failed.append(
                            result
                        )

                        print(
                            "STATUS: FAILED"
                        )

                except Exception as exc:
                    failure_dir = (
                        RECOVERY_FUNDS_OUTPUT_DIR
                        / (
                            f"{row}_"
                            f"{safe_filename(fund['excelPruAccessName'])}"
                            "_failed"
                        )
                    )

                    ensure_dir(
                        failure_dir
                    )

                    failure = {
                        "status": "failed",
                        "excelRow": row,
                        "prudentialUrl": fund[
                            "prudentialUrl"
                        ],
                        "excelPruAccessName": fund[
                            "excelPruAccessName"
                        ],
                        "error": str(
                            exc
                        ),
                        "recoveryStage": "Recovery 2",
                        "recovery1Failure": fund.get(
                            "recovery1Failure"
                        ),
                        "failedAtUtc": utc_now_iso(),
                    }

                    write_json(
                        failure_dir
                        / "failure.json",
                        failure,
                    )

                    failed.append(
                        failure
                    )

                    print(
                        "STATUS: FAILED"
                    )

                    print(
                        "Error:",
                        str(
                            exc
                        ),
                    )

        finally:
            context.close()
            browser.close()

    completed_at = utc_now_iso()

    summary = {
        "status": (
            "success"
            if not failed
            else "completed_with_failures"
        ),
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": str(
            EXCEL_FILE
        ),
        "recovery1RunSummary": str(
            RECOVERY_1_RUN_SUMMARY_FILE
        ),
        "recovery1FailureSource": failure_source,
        "recovery1ExcelFundUniverse": (
            recovery_1_summary.get(
                "excelFundUniverse"
            )
        ),
        "recovery1FailedFunds": len(
            recovery_1_failures
        ),
        "recovery1FailedRows": [
            int(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_1_failures
        ],
        "recovery2Universe": len(
            recovery_universe
        ),
        "recovery2UniverseRows": [
            int(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_universe
        ],
        "successfulFunds": len(
            successful
        ),
        "diagnosticFunds": len(
            diagnostic
        ),
        "failedFunds": len(
            failed
        ),
        "totalPublishedTopHoldings": len(
            all_holdings
        ),
        "successfulFundsDetail": successful,
        "diagnosticFundsDetail": diagnostic,
        "failedFundsDetail": failed,
        "rules": build_rules(),
    }

    write_json(
        RECOVERY_RUN_SUMMARY_FILE,
        summary,
    )

    write_json(
        RECOVERY_ALL_HOLDINGS_FILE,
        all_holdings,
    )

    print()
    print(
        "=" * 72
    )
    print(
        "RECOVERY 2 COMPLETE"
    )
    print(
        "=" * 72
    )
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
    print(
        "Recovery 2 universe:",
        len(
            recovery_universe
        ),
    )
    print(
        "Successful funds:",
        len(
            successful
        ),
    )
    print(
        "Diagnostic funds:",
        len(
            diagnostic
        ),
    )
    print(
        "Failed funds:",
        len(
            failed
        ),
    )
    print(
        "Total published holdings:",
        len(
            all_holdings
        ),
    )

    if diagnostic:
        print()
        print(
            "FUNDS WITHOUT STRICT CONFIRMATION:"
        )

        for item in diagnostic:
            print(
                f" - Row {item['excelRow']}: "
                f"{item.get('excelPruAccessName', '')}"
            )

    if failed:
        print()
        print(
            "FAILED FUNDS:"
        )

        for item in failed:
            print(
                f" - Row {item['excelRow']}: "
                f"{item.get('error', '')}"
            )

    print()
    print(
        "Output:",
        RECOVERY_OUTPUT_DIR,
    )
    print(
        "Run summary:",
        RECOVERY_RUN_SUMMARY_FILE,
    )
    print(
        "All holdings:",
        RECOVERY_ALL_HOLDINGS_FILE,
    )
    print(
        "=" * 72
    )

    return 0


# ============================================================================
# ENTRY POINT
# ============================================================================


if __name__ == "__main__":
    sys.exit(
        main()
    )
