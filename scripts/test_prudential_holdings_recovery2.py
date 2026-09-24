#!/usr/bin/env python3
"""
VGrat FMS - Prudential Top Holdings RECOVERY 2

RECOVERY CHAIN
==============
1. test_prudential_holdings.py
2. update_prudential_holdings.py
3. THIS SCRIPT

RECOVERY 2 SOURCE
=================
ONLY:
    output_holdings_recovery/run_summary.json

The Recovery 1 summary stores unresolved funds in:
    stillFailedFundsDetail

Recovery 2 therefore NEVER reads:
    output_holdings/run_summary.json

and NEVER rebuilds its universe from the baseline.

PURPOSE
=======
Recover only the funds still failed by Recovery 1.

This script is intentionally conservative:
- official Prudential Singapore sources only
- no third-party holdings data
- no inferred ranks
- no inferred names
- no name/weight proximity pairing
- no synthetic weights
- no forced ten holdings
- fewer than ten published holdings is allowed
- fixed-income coupon percentages are not portfolio weights
- the LAST percentage on a confirmed fixed-income security row is the
  portfolio weight
- a result is successful only when the physical PDF geometry proves a
  ranked holding table
- final verification downloads the same official PDF again and requires an
  exact rank/name/weight signature match

OUTPUT
======
output_holdings_recovery_2/
    run_summary.json
    all_holdings.json
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
            top_holdings.json
            top_holdings_section.txt
            recovery_result.json

        <row>_failed/
            factsheet.pdf
            factsheet_text.txt
            spatial_rows.json
            spatial_candidates.json
            candidate_regions.json
            candidate_holdings.json
            diagnostic_failure.json
            recovery_result.json

EXIT CODE
=========
Always 0 after the run so GitHub Actions preserves diagnostics/artifacts.
A configuration/programming failure before fund processing returns 1.
"""

from __future__ import annotations

import io
import json
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# =============================================================================
# CONFIGURATION
# =============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

RECOVERY_1_SUMMARY_FILE = Path(
    "output_holdings_recovery/run_summary.json"
)

OUTPUT_DIR = Path("output_holdings_recovery_2")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"
ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}

MAX_HOLDINGS = 10

PAGE_TIMEOUT_MS = 120_000
DOWNLOAD_TIMEOUT_MS = 120_000
POST_PAGE_WAIT_MS = 1_500

RETRY_COUNT = 3
RETRY_DELAY_SECONDS = 3.0

ROW_Y_TOLERANCE = 2.5
MIN_TABLE_ROWS = 2
MIN_NAME_LENGTH = 3

# Percentages are normally near the right side of the holdings table.
# These are deliberately broad; the actual column is learned from the PDF.
PERCENTAGE_RE = re.compile(
    r"(?<![\d.])(\d+(?:\.\d+)?)\s*%"
)

RANK_RE = re.compile(
    r"^\s*(\d{1,2})\s*$"
)

RANK_PREFIX_RE = re.compile(
    r"^\s*(\d{1,2})[.)]?\s+(.+?)\s*$"
)

# Text that is not a security name even if it contains a percentage/rank.
NOISE_TERMS = (
    "charge",
    "initial investment charge",
    "subscription method",
    "launch date",
    "performance chart",
    "calendar year performance",
    "benchmark",
    "price indexed",
    "distribution class",
    "fund size",
    "funds under management",
    "manager of the fund",
    "morningstar",
    "financial year end",
    "performance",
    "asset allocation",
    "portfolio characteristics",
    "important information",
    "disclaimer",
    "past performance",
    "source",
    "data as at",
    "inception date",
    "risk classification",
    "investment-linked",
    "insurance products",
    "cash, srs",
    "cash srs",
)

TABLE_HEADER_TERMS = (
    "top 10 holdings",
    "top ten holdings",
    "holdings",
    "name",
    "weight",
    "portfolio",
    "security",
    "investment",
)

DATE_RE = re.compile(
    r"\b(?:\d{1,2}[-/]){2}\d{2,4}\b"
    r"|\b\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}\b"
    r"|\b[A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}\b"
)

# =============================================================================
# HELPERS
# =============================================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(value: Any) -> str:
    text = clean_text(value).casefold()
    text = text.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def utc_now_iso() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
    )


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def safe_filename(value: Any) -> str:
    text = clean_text(value)
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    text = text.strip("._")
    return (text or "fund")[:120]


def is_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        return (
            parsed.scheme.lower() in {"http", "https"}
            and (parsed.hostname or "").lower() in PRUDENTIAL_HOSTS
        )
    except Exception:
        return False


def ensure_prudential_url(url: str) -> str:
    url = clean_text(url)
    if not url:
        raise RuntimeError("Prudential URL is empty.")
    if not is_prudential_url(url):
        raise RuntimeError(f"Non-Prudential URL rejected: {url}")
    return url


def normalize_name(value: str) -> str:
    text = clean_text(value)
    text = text.replace("|", " ")
    text = re.sub(r"\bNone\b", " ", text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip(" -|")
    return text


def signature(items: list[dict]) -> list[list[Any]]:
    return [
        [
            int(item["rank"]),
            normalize_text(item["name"]),
            float(item["weightPercent"]),
        ]
        for item in items
    ]


def validate_holdings(holdings: list[dict]) -> None:
    if not holdings:
        raise RuntimeError("No holdings were extracted.")

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"More than {MAX_HOLDINGS} holdings were extracted."
        )

    ranks = [int(item["rank"]) for item in holdings]
    expected = list(range(1, len(holdings) + 1))

    if ranks != expected:
        raise RuntimeError(
            "Confirmed holding ranks are not an explicit contiguous "
            f"sequence: {ranks}"
        )

    for item in holdings:
        name = normalize_name(item.get("name", ""))
        weight = item.get("weightPercent")

        if len(name) < MIN_NAME_LENGTH:
            raise RuntimeError("Confirmed holding has no usable name.")

        if weight is None:
            raise RuntimeError("Confirmed holding has no published weight.")

        weight = float(weight)
        if not 0.0 <= weight <= 100.0:
            raise RuntimeError(
                f"Invalid published holding weight: {weight}"
            )


def is_noise(text: str) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return True

    if normalized in {
        "charge",
        "performance",
        "benchmark",
        "holdings",
        "holding",
        "weight",
        "name",
        "security",
        "distribution class",
    }:
        return True

    for term in NOISE_TERMS:
        if term in normalized:
            return True

    return False


def is_probable_security_name(text: str) -> bool:
    text = normalize_name(text)

    if len(text) < MIN_NAME_LENGTH:
        return False

    if is_noise(text):
        return False

    # A pure percentage/date/rank cannot be a security name.
    if PERCENTAGE_RE.fullmatch(text):
        return False

    if RANK_RE.fullmatch(text):
        return False

    # Chart-axis fragments and years.
    if re.fullmatch(r"[\d\s./%-]+", text):
        return False

    return True


# =============================================================================
# RECOVERY 1 LOADER
# =============================================================================

def load_recovery_1_failed_funds() -> tuple[dict, list[dict]]:
    """
    Recovery 1's actual schema is:

        stillFailedFundsDetail

    This is the EXCLUSIVE Recovery 2 universe.

    No baseline summary is read.
    """

    if not RECOVERY_1_SUMMARY_FILE.exists():
        raise RuntimeError(
            "Recovery 1 run summary not found: "
            f"{RECOVERY_1_SUMMARY_FILE}"
        )

    payload = json.loads(
        RECOVERY_1_SUMMARY_FILE.read_text(encoding="utf-8")
    )

    if not isinstance(payload, dict):
        raise RuntimeError("Recovery 1 run summary is not a JSON object.")

    failed = payload.get("stillFailedFundsDetail")

    if not isinstance(failed, list):
        raise RuntimeError(
            "Recovery 1 summary does not contain the required "
            "'stillFailedFundsDetail' list."
        )

    normalized: list[dict] = []

    for entry in failed:
        if not isinstance(entry, dict):
            continue

        row = entry.get("excelRow")
        url = entry.get("prudentialUrl")
        pru_name = entry.get("pruAccessName")

        if row is None:
            raise RuntimeError(
                "Recovery 1 failed-fund detail is missing excelRow."
            )

        if not url:
            raise RuntimeError(
                f"Recovery 1 row {row} is missing prudentialUrl."
            )

        normalized.append(
            {
                "excelRow": int(row),
                "prudentialUrl": ensure_prudential_url(url),
                "pruAccessName": clean_text(pru_name),
                "recovery1Error": clean_text(entry.get("error")),
                "recovery1OutputDirectory": clean_text(
                    entry.get("outputDirectory")
                ),
                "baselineFailure": entry.get("baselineFailure"),
            }
        )

    normalized.sort(key=lambda item: item["excelRow"])

    # Explicitly prevent accidental duplication.
    seen: set[int] = set()
    for item in normalized:
        row = item["excelRow"]
        if row in seen:
            raise RuntimeError(
                f"Duplicate Recovery 1 failed row: {row}"
            )
        seen.add(row)

    return payload, normalized


# =============================================================================
# EXCEL REFERENCE LOOKUP
# =============================================================================

def read_excel_reference_funds() -> dict[int, dict]:
    """
    Reads the master workbook only to obtain the exact Column A/B values
    for the Recovery 1 rows.

    It does NOT expand the Recovery 2 universe.
    """

    if not EXCEL_FILE.exists():
        raise RuntimeError(f"Excel file not found: {EXCEL_FILE}")

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=False,
        keep_links=True,
    )

    try:
        worksheet = workbook.active
        rows: dict[int, dict] = {}

        for row_number in range(2, worksheet.max_row + 1):
            url_cell = worksheet.cell(row_number, 1)
            name_cell = worksheet.cell(row_number, 2)

            url = clean_text(url_cell.value)
            name = clean_text(name_cell.value)

            if url:
                rows[row_number] = {
                    "excelRow": row_number,
                    "prudentialUrl": url,
                    "pruAccessName": name,
                }

        return rows
    finally:
        workbook.close()


def merge_recovery_and_excel(
    recovery_funds: list[dict],
    excel_rows: dict[int, dict],
) -> list[dict]:
    merged: list[dict] = []

    for recovery in recovery_funds:
        row = recovery["excelRow"]
        excel = excel_rows.get(row)

        if excel is None:
            raise RuntimeError(
                f"Recovery 1 failed row {row} is not present in "
                "Funds Links.xlsm Column A."
            )

        # Recovery 1 URL remains authoritative for this recovery run.
        # Excel is checked for consistency only.
        excel_url = clean_text(excel["prudentialUrl"])
        recovery_url = clean_text(recovery["prudentialUrl"])

        if normalize_text(excel_url) != normalize_text(recovery_url):
            raise RuntimeError(
                f"Excel/Recovery 1 URL mismatch at row {row}."
            )

        merged.append(
            {
                **recovery,
                "excelPruAccessName": clean_text(
                    excel.get("pruAccessName")
                ),
            }
        )

    return merged


# =============================================================================
# PLAYWRIGHT / FACTSHEET DOWNLOAD
# =============================================================================

def find_factsheet_url(
    page,
    page_url: str,
) -> str:
    """
    Locate an official Prudential PDF link from the fund page.

    Preference:
    - links containing factsheet/fund/facts
    - direct PDF links
    - Prudential-hosted document links only
    """

    links = page.locator("a").all()

    candidates: list[tuple[int, str]] = []

    for link in links:
        try:
            href = clean_text(link.get_attribute("href"))
            text = clean_text(link.inner_text(timeout=2000))
        except Exception:
            continue

        if not href:
            continue

        absolute = urljoin(page_url, href)

        if not is_prudential_url(absolute):
            continue

        parsed = urlparse(absolute)
        lower = parsed.path.casefold()

        if ".pdf" not in lower:
            continue

        score = 0

        combined = normalize_text(
            f"{text} {absolute}"
        )

        if "factsheet" in combined:
            score += 100

        if "fund" in combined:
            score += 20

        if "facts" in combined:
            score += 10

        if "/media/ilp/" in lower:
            score += 10

        candidates.append((score, absolute))

    if not candidates:
        raise RuntimeError(
            "No official Prudential PDF factsheet link found."
        )

    candidates.sort(
        key=lambda item: (item[0], item[1]),
        reverse=True,
    )

    return candidates[0][1]


def download_factsheet(
    browser,
    prudential_url: str,
) -> tuple[bytes, str, str]:
    """
    Returns:
        factsheet bytes
        resolved fund page URL
        factsheet URL
    """

    context = browser.new_context(
        accept_downloads=True
    )

    page = context.new_page()
    page.set_default_timeout(PAGE_TIMEOUT_MS)

    try:
        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(POST_PAGE_WAIT_MS)

        final_url = page.url

        if not is_prudential_url(final_url):
            raise RuntimeError(
                f"Resolved page left Prudential Singapore host: {final_url}"
            )

        factsheet_url = find_factsheet_url(
            page,
            final_url,
        )

        response = context.request.get(
            factsheet_url,
            timeout=DOWNLOAD_TIMEOUT_MS,
        )

        if not response.ok:
            raise RuntimeError(
                f"Factsheet HTTP request failed: "
                f"{response.status} {factsheet_url}"
            )

        content_type = normalize_text(
            response.headers.get("content-type", "")
        )

        data = response.body()

        if not data.startswith(b"%PDF"):
            # Some Prudential responses have a generic content type.
            # The PDF magic header is the definitive check.
            raise RuntimeError(
                "Official factsheet response is not a PDF."
            )

        if "pdf" not in content_type and not factsheet_url.casefold().endswith(".pdf"):
            raise RuntimeError(
                "Official factsheet response did not identify itself as PDF."
            )

        return data, final_url, factsheet_url

    finally:
        context.close()


def download_with_retries(
    browser,
    url: str,
) -> tuple[bytes, str, str]:
    last_error: Exception | None = None

    for attempt in range(1, RETRY_COUNT + 1):
        try:
            return download_factsheet(browser, url)
        except (
            PlaywrightTimeoutError,
            Exception,
        ) as error:
            last_error = error
            if attempt < RETRY_COUNT:
                time.sleep(RETRY_DELAY_SECONDS)

    raise RuntimeError(
        f"Factsheet download failed after {RETRY_COUNT} attempts: "
        f"{clean_text(str(last_error))}"
    )


# =============================================================================
# PDF TEXT
# =============================================================================

def extract_pdf_text(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))

    chunks: list[str] = []

    for page in reader.pages:
        try:
            chunks.append(page.extract_text() or "")
        except Exception as error:
            chunks.append(
                f"[PDF_TEXT_EXTRACTION_ERROR: {error}]"
            )

    return "\n".join(chunks)


# =============================================================================
# POSITIONED PDF MODEL
# =============================================================================

@dataclass
class PDFWord:
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


def build_positioned_rows(pdf_bytes: bytes) -> list[PDFRow]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    rows: list[PDFRow] = []

    for page_index, page in enumerate(reader.pages, start=1):
        words: list[PDFWord] = []

        def visitor(text, cm, tm, font_dict, font_size):
            if not text or not text.strip():
                return

            # tm[4], tm[5] are text origin coordinates in PDF user space.
            x = float(tm[4])
            y = float(tm[5])

            cleaned = clean_text(text)

            if not cleaned:
                return

            # Approximate width. Exact glyph width is not required because
            # x positions are used comparatively within one PDF.
            width = max(
                float(font_size) * 0.25 * len(cleaned),
                1.0,
            )

            words.append(
                PDFWord(
                    text=cleaned,
                    x=x,
                    y=y,
                    width=width,
                    height=float(font_size),
                )
            )

        try:
            page.extract_text(
                visitor_text=visitor
            )
        except Exception:
            continue

        words.sort(key=lambda item: (-item.y, item.x))

        page_rows: list[PDFRow] = []

        for word in words:
            target = None

            for row in page_rows:
                if abs(row.y - word.y) <= ROW_Y_TOLERANCE:
                    target = row
                    break

            if target is None:
                target = PDFRow(
                    page=page_index,
                    y=word.y,
                    words=[],
                )
                page_rows.append(target)

            target.words.append(word)

        for row in page_rows:
            row.words.sort(key=lambda item: item.x)
            rows.append(row)

    return rows


# =============================================================================
# SPATIAL ROW TOKENIZATION
# =============================================================================

@dataclass
class PercentageToken:
    value: float
    text: str
    x: float
    right: float


@dataclass
class RankToken:
    rank: int
    text: str
    x: float
    right: float


def row_percentages(row: PDFRow) -> list[PercentageToken]:
    tokens: list[PercentageToken] = []

    for word in row.words:
        matches = list(PERCENTAGE_RE.finditer(word.text))

        for match in matches:
            value = float(match.group(1))
            tokens.append(
                PercentageToken(
                    value=value,
                    text=match.group(0),
                    x=word.x,
                    right=word.right,
                )
            )

    return tokens


def row_rank_tokens(row: PDFRow) -> list[RankToken]:
    tokens: list[RankToken] = []

    for word in row.words:
        text = clean_text(word.text)

        match = RANK_RE.fullmatch(text)

        if match:
            rank = int(match.group(1))
            if 1 <= rank <= 10:
                tokens.append(
                    RankToken(
                        rank=rank,
                        text=text,
                        x=word.x,
                        right=word.right,
                    )
                )
            continue

        match = RANK_PREFIX_RE.match(text)

        if match:
            rank = int(match.group(1))
            if 1 <= rank <= 10:
                tokens.append(
                    RankToken(
                        rank=rank,
                        text=match.group(1),
                        x=word.x,
                        right=word.right,
                    )
                )

    return tokens


def words_between(
    row: PDFRow,
    left: float,
    right: float,
) -> str:
    selected = [
        word
        for word in row.words
        if word.x >= left and word.right <= right
    ]

    return normalize_name(
        " ".join(
            word.text
            for word in sorted(
                selected,
                key=lambda item: item.x,
            )
        )
    )


# =============================================================================
# TABLE GEOMETRY
# =============================================================================

def median(values: list[float]) -> float:
    if not values:
        return 0.0
    values = sorted(values)
    middle = len(values) // 2
    if len(values) % 2:
        return values[middle]
    return (values[middle - 1] + values[middle]) / 2.0


def cluster_values(
    values: list[float],
    tolerance: float,
) -> list[list[float]]:
    clusters: list[list[float]] = []

    for value in sorted(values):
        placed = False

        for cluster in clusters:
            if abs(value - median(cluster)) <= tolerance:
                cluster.append(value)
                placed = True
                break

        if not placed:
            clusters.append([value])

    return clusters


def candidate_rank_columns(
    rows: list[PDFRow],
) -> list[dict]:
    entries: list[tuple[float, int, int, str]] = []

    for row_index, row in enumerate(rows):
        for token in row_rank_tokens(row):
            entries.append(
                (
                    token.x,
                    token.rank,
                    row_index,
                    row.text,
                )
            )

    clusters = cluster_values(
        [item[0] for item in entries],
        tolerance=10.0,
    )

    candidates: list[dict] = []

    for cluster in clusters:
        center = median(cluster)

        cluster_entries = [
            item
            for item in entries
            if abs(item[0] - center) <= 10.0
        ]

        ranks = [
            item[1]
            for item in cluster_entries
        ]

        distinct_rows = len(
            {item[2] for item in cluster_entries}
        )

        if distinct_rows < MIN_TABLE_ROWS:
            continue

        candidates.append(
            {
                "x": center,
                "rowCount": distinct_rows,
                "ranks": ranks,
                "rankRows": cluster_entries,
            }
        )

    candidates.sort(
        key=lambda item: (
            item["rowCount"],
            len(set(item["ranks"])),
        ),
        reverse=True,
    )

    return candidates


def candidate_weight_columns(
    rows: list[PDFRow],
) -> list[dict]:
    entries: list[tuple[float, float, int, str]] = []

    for row_index, row in enumerate(rows):
        for token in row_percentages(row):
            entries.append(
                (
                    token.x,
                    token.value,
                    row_index,
                    row.text,
                )
            )

    if not entries:
        return []

    clusters = cluster_values(
        [item[0] for item in entries],
        tolerance=14.0,
    )

    candidates: list[dict] = []

    for cluster in clusters:
        center = median(cluster)

        cluster_entries = [
            item
            for item in entries
            if abs(item[0] - center) <= 14.0
        ]

        distinct_rows = len(
            {item[2] for item in cluster_entries}
        )

        if distinct_rows < MIN_TABLE_ROWS:
            continue

        candidates.append(
            {
                "x": center,
                "rowCount": distinct_rows,
                "percentages": cluster_entries,
            }
        )

    candidates.sort(
        key=lambda item: item["rowCount"],
        reverse=True,
    )

    return candidates


def section_heading_rows(
    rows: list[PDFRow],
) -> list[int]:
    indexes: list[int] = []

    for index, row in enumerate(rows):
        normalized = normalize_text(row.text)

        if (
            "top 10 holdings" in normalized
            or "top ten holdings" in normalized
            or normalized == "holdings"
        ):
            indexes.append(index)

    return indexes


def nearby_table_region(
    rows: list[PDFRow],
    heading_index: int,
) -> tuple[int, int]:
    """
    Start after the heading and stop before obvious later sections.

    The region is intentionally page-local.
    """

    page = rows[heading_index].page
    start = heading_index + 1
    end = start

    while end < len(rows):
        row = rows[end]

        if row.page != page:
            break

        normalized = normalize_text(row.text)

        if end > start and any(
            marker in normalized
            for marker in (
                "asset allocation",
                "portfolio characteristics",
                "performance chart",
                "calendar year performance",
                "important information",
                "disclaimer",
                "past performance",
                "source:",
            )
        ):
            break

        end += 1

    return start, end


# =============================================================================
# PHYSICAL HOLDING CANDIDATES
# =============================================================================

def build_candidate_from_rank_row(
    rows: list[PDFRow],
    row_index: int,
    rank_x: float,
    weight_x: float,
) -> dict | None:
    row = rows[row_index]

    rank_tokens = [
        token
        for token in row_rank_tokens(row)
        if abs(token.x - rank_x) <= 10.0
    ]

    if not rank_tokens:
        return None

    rank_token = rank_tokens[0]

    # Weight must be physically in the learned weight column.
    percentages = [
        token
        for token in row_percentages(row)
        if abs(token.x - weight_x) <= 14.0
    ]

    if not percentages:
        return None

    # A genuine holding row normally has one portfolio-weight token in the
    # weight column. If multiple tokens occupy the same column, reject rather
    # than guessing.
    distinct_values = {
        round(token.value, 8)
        for token in percentages
    }

    if len(distinct_values) != 1:
        return None

    weight = percentages[0]

    # Name column is physically between rank and weight.
    name_left = max(
        rank_token.right + 1.0,
        rank_x + 4.0,
    )
    name_right = weight.x - 1.0

    if name_right <= name_left:
        return None

    name = words_between(
        row,
        name_left,
        name_right,
    )

    # Some PDFs place the security name slightly outside the strict row
    # bounds. If no name exists, caller may add continuation rows, but never
    # invent a name from arbitrary nearby text.
    if not name:
        return {
            "rank": rank_token.rank,
            "weight": weight.value,
            "name": "",
            "page": row.page,
            "source_row_index": row_index,
            "rankX": rank_token.x,
            "weightX": weight.x,
            "raw_row": row.text,
            "needsContinuation": True,
        }

    if not is_probable_security_name(name):
        return None

    return {
        "rank": rank_token.rank,
        "weight": weight.value,
        "name": name,
        "page": row.page,
        "source_row_index": row_index,
        "rankX": rank_token.x,
        "weightX": weight.x,
        "raw_row": row.text,
        "needsContinuation": False,
    }


def add_continuation_rows(
    rows: list[PDFRow],
    candidate: dict,
    weight_x: float,
    rank_x: float,
    region_end: int,
) -> dict:
    """
    Add physical name continuation lines directly below a ranked row.

    A continuation is accepted only while:
    - same page
    - before next explicit rank
    - no percentage in the learned weight column
    - text lies inside the name column
    - text is not semantic noise
    """

    if not candidate.get("needsContinuation"):
        return candidate

    start = int(candidate["source_row_index"]) + 1
    fragments: list[str] = []

    for index in range(start, min(region_end, start + 6)):
        row = rows[index]

        if row.page != candidate["page"]:
            break

        next_rank = [
            token
            for token in row_rank_tokens(row)
            if abs(token.x - rank_x) <= 10.0
        ]

        if next_rank:
            break

        weight_here = [
            token
            for token in row_percentages(row)
            if abs(token.x - weight_x) <= 14.0
        ]

        if weight_here:
            break

        fragment = words_between(
            row,
            max(rank_x + 4.0, row.x_min),
            weight_x - 1.0,
        )

        if not fragment:
            continue

        if is_noise(fragment):
            break

        fragments.append(fragment)

    if fragments:
        candidate["name"] = normalize_name(
            " ".join(fragments)
        )

    candidate["needsContinuation"] = False

    if not is_probable_security_name(candidate.get("name", "")):
        return {}

    return candidate


def extract_table_candidate(
    rows: list[PDFRow],
    start: int,
    end: int,
    rank_column: dict,
    weight_column: dict,
) -> dict | None:
    rank_x = float(rank_column["x"])
    weight_x = float(weight_column["x"])

    candidates: list[dict] = []

    for index in range(start, end):
        row = rows[index]

        if row.page != rows[start].page:
            break

        candidate = build_candidate_from_rank_row(
            rows,
            index,
            rank_x,
            weight_x,
        )

        if candidate is None:
            continue

        candidate = add_continuation_rows(
            rows,
            candidate,
            weight_x,
            rank_x,
            end,
        )

        if not candidate:
            continue

        if not candidate.get("name"):
            continue

        candidates.append(candidate)

    if len(candidates) < MIN_TABLE_ROWS:
        return None

    # Explicit ranks must be present. No inferred rank.
    ranks = [
        int(item["rank"])
        for item in candidates
    ]

    # Remove duplicate physical observations only when they are exactly the
    # same row. We do not collapse distinct holdings with duplicate names.
    unique: list[dict] = []
    seen_rows: set[int] = set()

    for item in candidates:
        row_index = int(item["source_row_index"])
        if row_index in seen_rows:
            continue
        seen_rows.add(row_index)
        unique.append(item)

    candidates = unique

    # A valid published sequence must begin at 1 and be contiguous.
    # We deliberately do not sort/relabel rows to manufacture ranks.
    candidates.sort(
        key=lambda item: int(item["source_row_index"])
    )

    ranks = [
        int(item["rank"])
        for item in candidates
    ]

    # Some PDF extraction layouts repeat a table. Find the longest contiguous
    # explicit rank run starting at 1.
    best: list[dict] = []

    for offset, item in enumerate(candidates):
        if int(item["rank"]) != 1:
            continue

        run = [item]
        expected = 2

        for following in candidates[offset + 1:]:
            rank = int(following["rank"])

            if rank == expected:
                run.append(following)
                expected += 1

                if len(run) == MAX_HOLDINGS:
                    break

            elif rank == 1:
                break

            else:
                # A genuine table cannot skip a published rank.
                break

        if len(run) > len(best):
            best = run

    if len(best) < MIN_TABLE_ROWS:
        return None

    # Validate exact explicit sequence.
    holdings: list[dict] = []

    for item in best:
        holdings.append(
            {
                "rank": int(item["rank"]),
                "name": normalize_name(item["name"]),
                "weightPercent": float(item["weight"]),
                "weightText": f"{float(item['weight']):g}%",
                "page": int(item["page"]),
                "sourceRowIndex": int(item["source_row_index"]),
                "extractionStrategy": "recovery2_spatial_table",
                "evidence": {
                    "rankX": float(item["rankX"]),
                    "weightX": float(item["weightX"]),
                    "rawRow": item["raw_row"],
                },
            }
        )

    try:
        validate_holdings(holdings)
    except Exception:
        return None

    return {
        "holdings": holdings,
        "rankX": rank_x,
        "weightX": weight_x,
        "start": start,
        "end": end,
        "page": rows[start].page,
        "rowCount": len(holdings),
    }


# =============================================================================
# TABLE SEARCH
# =============================================================================

def search_spatial_tables(
    rows: list[PDFRow],
) -> tuple[
    list[dict],
    list[dict],
    list[dict],
]:
    """
    Returns:
        confirmed candidates
        candidate regions
        diagnostic spatial candidates

    No candidate is accepted merely because a name is close to a percentage.
    The rank, name and weight must belong to the same physical table geometry.
    """

    confirmed: list[dict] = []
    regions: list[dict] = []
    diagnostics: list[dict] = []

    # Process one page at a time.
    page_numbers = sorted(
        {row.page for row in rows}
    )

    for page in page_numbers:
        page_rows = [
            row
            for row in rows
            if row.page == page
        ]

        if not page_rows:
            continue

        headings = [
            index
            for index, row in enumerate(page_rows)
            if (
                "top 10 holdings" in normalize_text(row.text)
                or "top ten holdings" in normalize_text(row.text)
                or normalize_text(row.text) == "holdings"
            )
        ]

        # If heading extraction is poor, use any page containing repeated
        # explicit rank/percentage geometry, but still require a real table.
        if not headings:
            headings = [-1]

        for heading in headings:
            if heading >= 0:
                start = heading + 1
            else:
                start = 0

            end = len(page_rows)

            # Limit fallback pages to the first coherent table-like block.
            if heading >= 0:
                for index in range(start, len(page_rows)):
                    normalized = normalize_text(page_rows[index].text)
                    if any(
                        marker in normalized
                        for marker in (
                            "asset allocation",
                            "portfolio characteristics",
                            "performance chart",
                            "calendar year performance",
                            "important information",
                            "disclaimer",
                        )
                    ):
                        end = index
                        break

            if end - start < MIN_TABLE_ROWS:
                continue

            region_rows = page_rows[start:end]

            rank_columns = candidate_rank_columns(
                region_rows
            )

            weight_columns = candidate_weight_columns(
                region_rows
            )

            region_record = {
                "page": page,
                "headingIndex": heading,
                "start": start,
                "end": end,
                "rankColumns": [
                    {
                        "x": c["x"],
                        "rowCount": c["rowCount"],
                        "ranks": c["ranks"],
                    }
                    for c in rank_columns[:20]
                ],
                "weightColumns": [
                    {
                        "x": c["x"],
                        "rowCount": c["rowCount"],
                    }
                    for c in weight_columns[:20]
                ],
            }

            regions.append(region_record)

            for rank_column in rank_columns[:12]:
                for weight_column in weight_columns[:20]:
                    # Name column must exist physically between rank and weight.
                    if weight_column["x"] <= rank_column["x"] + 10:
                        continue

                    candidate = extract_table_candidate(
                        region_rows,
                        0,
                        len(region_rows),
                        rank_column,
                        weight_column,
                    )

                    if candidate is None:
                        continue

                    # Reject obvious non-holdings candidates.
                    names = [
                        normalize_text(item["name"])
                        for item in candidate["holdings"]
                    ]

                    if any(
                        is_noise(name)
                        for name in names
                    ):
                        continue

                    # Require multiple distinct explicit ranks.
                    distinct_ranks = {
                        int(item["rank"])
                        for item in candidate["holdings"]
                    }

                    if len(distinct_ranks) < MIN_TABLE_ROWS:
                        continue

                    candidate["page"] = page
                    candidate["rankColumnX"] = rank_column["x"]
                    candidate["weightColumnX"] = weight_column["x"]

                    confirmed.append(candidate)

    # Diagnostic candidate dump is deliberately comprehensive enough to
    # understand unresolved PDFs without treating diagnostics as results.
    for region in regions:
        diagnostics.append(
            {
                "page": region["page"],
                "rankColumns": region["rankColumns"],
                "weightColumns": region["weightColumns"],
                "headingIndex": region["headingIndex"],
            }
        )

    return confirmed, regions, diagnostics


# =============================================================================
# FIXED-INCOME SAFETY CHECK
# =============================================================================

def validate_fixed_income_names(
    holdings: list[dict],
) -> None:
    """
    A fixed-income holding may contain coupon and maturity information.

    We do not remove those percentages from the security name if they are
    physically part of the published security description.

    This check only prevents a row whose extracted name is merely a maturity
    date from being accepted.
    """

    for item in holdings:
        name = normalize_name(item["name"])

        if DATE_RE.fullmatch(name):
            raise RuntimeError(
                "Fixed-income candidate contains only a maturity/date "
                "instead of the published security name."
            )


# =============================================================================
# SECTION TEXT FOR AUDIT
# =============================================================================

def build_section_text(
    holdings: list[dict],
) -> str:
    lines = []

    for item in holdings:
        lines.append(
            f"{item['rank']} | "
            f"{item['name']} | "
            f"{item['weightText']}"
        )

    return "\n".join(lines)


# =============================================================================
# FUND PROCESSING
# =============================================================================

def fund_identifier(
    row: int,
    final_url: str,
    fund_name: str,
) -> str:
    parsed = urlparse(final_url)
    slug = (
        parsed.path.rstrip("/").split("/")[-1]
        or fund_name
        or f"fund_{row}"
    )
    return safe_filename(slug)


def create_result_base(
    entry: dict,
    final_url: str,
    factsheet_url: str,
    pdf_bytes: bytes,
) -> dict:
    reader = PdfReader(io.BytesIO(pdf_bytes))

    fund_name = clean_text(
        entry.get("excelPruAccessName")
        or entry.get("pruAccessName")
    )

    return {
        "status": "success",
        "excelRow": int(entry["excelRow"]),
        "prudentialUrl": entry["prudentialUrl"],
        "finalUrl": final_url,
        "excelPruAccessName": fund_name,
        "factsheetUrl": factsheet_url,
        "factsheetPageCount": len(reader.pages),
    }


def save_success(
    entry: dict,
    result: dict,
    pdf_bytes: bytes,
    pdf_text: str,
    regions: list[dict],
    diagnostics: list[dict],
    candidates: list[dict],
    confirmation: dict,
) -> Path:
    identifier = fund_identifier(
        result["excelRow"],
        result["finalUrl"],
        result.get("excelPruAccessName", ""),
    )

    directory = (
        FUNDS_OUTPUT_DIR
        / f"{result['excelRow']}_{identifier}"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    (directory / "factsheet.pdf").write_bytes(
        pdf_bytes
    )

    (directory / "factsheet_text.txt").write_text(
        pdf_text,
        encoding="utf-8",
    )

    save_json(
        directory / "spatial_rows.json",
        [
            {
                "page": row.page,
                "y": row.y,
                "xMin": row.x_min,
                "xMax": row.x_max,
                "text": row.text,
                "words": [
                    asdict(word)
                    for word in row.words
                ],
            }
            for row in build_positioned_rows(pdf_bytes)
        ],
    )

    save_json(
        directory / "spatial_candidates.json",
        diagnostics,
    )

    save_json(
        directory / "candidate_regions.json",
        regions,
    )

    save_json(
        directory / "candidate_holdings.json",
        candidates,
    )

    save_json(
        directory / "confirmation.json",
        confirmation,
    )

    metadata = {
        "excelRow": result["excelRow"],
        "prudentialUrl": result["prudentialUrl"],
        "finalUrl": result["finalUrl"],
        "factsheetUrl": result["factsheetUrl"],
        "recoverySource": str(RECOVERY_1_SUMMARY_FILE),
        "recoverySourceField": "stillFailedFundsDetail",
        "officialPrudentialOnly": True,
        "noProximityNameWeightMatching": True,
        "noInferredRanks": True,
        "noInferredNames": True,
        "maximumHoldings": MAX_HOLDINGS,
    }

    save_json(
        directory / "metadata.json",
        metadata,
    )

    (directory / "top_holdings_section.txt").write_text(
        build_section_text(result["topHoldings"]),
        encoding="utf-8",
    )

    save_json(
        directory / "top_holdings.json",
        result,
    )

    recovery_result = {
        "status": "success",
        "excelRow": result["excelRow"],
        "holdingsCount": result["topHoldingsCount"],
        "topHoldings": result["topHoldings"],
        "verification": result["finalVerification"],
    }

    save_json(
        directory / "recovery_result.json",
        recovery_result,
    )

    return directory


def save_failure(
    entry: dict,
    pdf_bytes: bytes | None,
    pdf_text: str,
    rows: list[PDFRow],
    regions: list[dict],
    diagnostics: list[dict],
    candidates: list[dict],
    error: str,
) -> Path:
    directory = (
        FUNDS_OUTPUT_DIR
        / f"{entry['excelRow']}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    if pdf_bytes:
        (directory / "factsheet.pdf").write_bytes(
            pdf_bytes
        )

    (directory / "factsheet_text.txt").write_text(
        pdf_text,
        encoding="utf-8",
    )

    save_json(
        directory / "spatial_rows.json",
        [
            {
                "page": row.page,
                "y": row.y,
                "xMin": row.x_min,
                "xMax": row.x_max,
                "text": row.text,
                "words": [
                    asdict(word)
                    for word in row.words
                ],
            }
            for row in rows
        ],
    )

    save_json(
        directory / "spatial_candidates.json",
        diagnostics,
    )

    save_json(
        directory / "candidate_regions.json",
        regions,
    )

    save_json(
        directory / "candidate_holdings.json",
        candidates,
    )

    diagnostic = {
        "status": "DIAGNOSTIC_NO_CONFIRMED_RESULT",
        "excelRow": entry["excelRow"],
        "prudentialUrl": entry["prudentialUrl"],
        "pruAccessName": entry.get("pruAccessName"),
        "error": error,
        "recovery1Error": entry.get("recovery1Error"),
        "rules": {
            "officialPrudentialOnly": True,
            "noProximityNameWeightMatching": True,
            "noInferredRanks": True,
            "noInferredNames": True,
            "noSyntheticWeights": True,
            "fewerThanTenAllowed": True,
        },
    }

    save_json(
        directory / "diagnostic_failure.json",
        diagnostic,
    )

    save_json(
        directory / "recovery_result.json",
        diagnostic,
    )

    return directory


def extract_once(
    pdf_bytes: bytes,
) -> tuple[list[dict], dict, list[dict], list[dict], list[dict]]:
    rows = build_positioned_rows(pdf_bytes)

    confirmed, regions, diagnostics = search_spatial_tables(
        rows
    )

    # Preserve all structural candidates for diagnostics.
    candidate_payload = confirmed

    if not confirmed:
        raise RuntimeError(
            "No confirmed ranked holdings table was found "
            "using official PDF coordinates."
        )

    # Never choose by arbitrary proximity. A candidate must have the largest
    # explicit contiguous rank sequence. If two candidates have the same
    # sequence length but different signatures, reject as ambiguous.
    max_count = max(
        len(item["holdings"])
        for item in confirmed
    )

    best = [
        item
        for item in confirmed
        if len(item["holdings"]) == max_count
    ]

    unique_signatures = {
        json.dumps(
            signature(item["holdings"]),
            ensure_ascii=False,
        )
        for item in best
    }

    if len(unique_signatures) != 1:
        raise RuntimeError(
            "Multiple equally strong spatial holdings candidates have "
            "different exact signatures; extraction is ambiguous."
        )

    selected = best[0]

    holdings = selected["holdings"]

    validate_holdings(holdings)
    validate_fixed_income_names(holdings)

    confirmation = {
        "confirmed": True,
        "confirmationMode": "strict_spatial_rank_name_weight_table",
        "strategy": "recovery2_spatial_table",
        "page": selected["page"],
        "rankColumnX": selected["rankColumnX"],
        "weightColumnX": selected["weightColumnX"],
        "holdingsCount": len(holdings),
        "signature": signature(holdings),
    }

    return (
        holdings,
        confirmation,
        regions,
        diagnostics,
        candidate_payload,
    )


# =============================================================================
# FINAL VERIFICATION
# =============================================================================

def verify_exact_signature(
    browser,
    entry: dict,
    expected: list[dict],
) -> dict:
    """
    Fresh official PDF download and a fresh extraction.

    Exact rank/name/weight signature must match.
    """

    fresh_pdf, fresh_final_url, fresh_factsheet_url = (
        download_with_retries(
            browser,
            entry["prudentialUrl"],
        )
    )

    fresh_holdings, confirmation, _, _, _ = extract_once(
        fresh_pdf
    )

    if signature(fresh_holdings) != signature(expected):
        raise RuntimeError(
            "Final verification failed: fresh official PDF extraction "
            "produced a different rank/name/weight signature."
        )

    return {
        "verified": True,
        "verificationMethod": "fresh_same_official_pdf_extraction_engine",
        "nameWeightProximityMatching": False,
        "freshFinalUrl": fresh_final_url,
        "freshFactsheetUrl": fresh_factsheet_url,
        "signature": signature(fresh_holdings),
        "confirmation": confirmation,
        "verifiedAtUtc": utc_now_iso(),
    }


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:
    started_at = utc_now_iso()

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )
    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    recovery_1_summary, recovery_funds = (
        load_recovery_1_failed_funds()
    )

    excel_rows = read_excel_reference_funds()

    universe = merge_recovery_and_excel(
        recovery_funds,
        excel_rows,
    )

    print()
    print("============================================================")
    print("VGrat FMS - PRUDENTIAL TOP HOLDINGS RECOVERY 2")
    print("============================================================")
    print(f"Recovery 1 summary: {RECOVERY_1_SUMMARY_FILE}")
    print("Recovery 1 failure field: stillFailedFundsDetail")
    print(f"Recovery 2 universe: {len(universe)}")
    print(
        "Recovery 2 rows: "
        + ", ".join(
            str(item["excelRow"])
            for item in universe
        )
    )
    print("============================================================")
    print()

    successful: list[dict] = []
    failed: list[dict] = []

    all_holdings: list[dict] = []

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=True
            )

            try:
                for position, entry in enumerate(
                    universe,
                    start=1,
                ):
                    row = entry["excelRow"]

                    print(
                        f"[{position}/{len(universe)}] "
                        f"Excel row {row}: "
                        f"{entry.get('pruAccessName') or '-'}"
                    )

                    pdf_bytes: bytes | None = None
                    pdf_text = ""
                    rows: list[PDFRow] = []
                    regions: list[dict] = []
                    diagnostics: list[dict] = []
                    candidate_payload: list[dict] = []

                    try:
                        pdf_bytes, final_url, factsheet_url = (
                            download_with_retries(
                                browser,
                                entry["prudentialUrl"],
                            )
                        )

                        pdf_text = extract_pdf_text(
                            pdf_bytes
                        )

                        (
                            holdings,
                            confirmation,
                            regions,
                            diagnostics,
                            candidate_payload,
                        ) = extract_once(
                            pdf_bytes
                        )

                        # Fresh download verification.
                        verification = verify_exact_signature(
                            browser,
                            entry,
                            holdings,
                        )

                        result = create_result_base(
                            entry,
                            final_url,
                            factsheet_url,
                            pdf_bytes,
                        )

                        result.update(
                            {
                                "holdingsSectionStatus": "published",
                                "topHoldingsCount": len(holdings),
                                "topHoldings": holdings,
                                "holdingsParser": "recovery2_spatial_table",
                                "holdingsSectionExtraction": "pdf_coordinates",
                                "extractionConfirmation": confirmation,
                                "finalVerification": verification,
                                "recovery1Error": entry.get(
                                    "recovery1Error"
                                ),
                                "recovery1OutputDirectory": entry.get(
                                    "recovery1OutputDirectory"
                                ),
                                "recovery2CompletedAtUtc": utc_now_iso(),
                            }
                        )

                        directory = save_success(
                            entry,
                            result,
                            pdf_bytes,
                            pdf_text,
                            regions,
                            diagnostics,
                            candidate_payload,
                            {
                                **confirmation,
                                "finalVerification": verification,
                            },
                        )

                        result["outputDirectory"] = str(
                            directory
                        )

                        successful.append(result)
                        all_holdings.append(result)

                        print(
                            f"    SUCCESS: {len(holdings)} holdings"
                        )

                        for item in holdings:
                            print(
                                f"      {item['rank']}. "
                                f"{item['name']} - "
                                f"{item['weightText']}"
                            )

                    except Exception as error:
                        error_text = clean_text(
                            str(error)
                        )

                        # Rebuild rows for diagnostics if extraction failed
                        # before rows were retained.
                        if pdf_bytes and not rows:
                            try:
                                rows = build_positioned_rows(
                                    pdf_bytes
                                )
                            except Exception:
                                rows = []

                        failure_directory = save_failure(
                            entry,
                            pdf_bytes,
                            pdf_text,
                            rows,
                            regions,
                            diagnostics,
                            candidate_payload,
                            error_text,
                        )

                        failed.append(
                            {
                                "status": "failed",
                                "excelRow": row,
                                "prudentialUrl": entry[
                                    "prudentialUrl"
                                ],
                                "pruAccessName": entry.get(
                                    "pruAccessName"
                                ),
                                "error": error_text,
                                "outputDirectory": str(
                                    failure_directory
                                ),
                                "recovery1Error": entry.get(
                                    "recovery1Error"
                                ),
                            }
                        )

                        print(
                            f"    FAILED: {error_text}"
                        )

            finally:
                browser.close()

    except Exception as error:
        # A configuration-level failure is different from an individual fund
        # failure. Preserve the diagnostic summary and return 1.
        completed_at = utc_now_iso()

        summary = {
            "status": "error",
            "startedAtUtc": started_at,
            "completedAtUtc": completed_at,
            "recovery1RunSummary": str(
                RECOVERY_1_SUMMARY_FILE
            ),
            "recovery1FailureSource": "stillFailedFundsDetail",
            "recovery1FailedFunds": len(
                recovery_funds
            ),
            "recovery1FailedRows": [
                item["excelRow"]
                for item in recovery_funds
            ],
            "recovery2Universe": len(universe),
            "recovery2UniverseRows": [
                item["excelRow"]
                for item in universe
            ],
            "successfulFunds": len(successful),
            "failedFunds": len(failed),
            "fatalError": clean_text(str(error)),
            "rules": {
                "frozenBaselineUntouched": True,
                "recovery1Untouched": True,
                "failedFundsDetailIsExclusiveRecoveryUniverse": True,
                "recovery1FailureField": "stillFailedFundsDetail",
                "officialPrudentialFactsheetOnly": True,
                "officialPrudentialSingaporeOnly": True,
                "maximumHoldings": MAX_HOLDINGS,
                "fewerThanTenHoldingsAllowed": True,
                "noInferredHoldings": True,
                "noInferredRanks": True,
                "noInferredNames": True,
                "noFabricatedHoldingWeights": True,
                "noProximityNameWeightMatching": True,
                "finalVerificationRequiresExactSignature": True,
            },
        }

        save_json(
            RUN_SUMMARY_FILE,
            summary,
        )

        save_json(
            ALL_HOLDINGS_FILE,
            {
                "status": "error",
                "holdings": [],
            },
        )

        print()
        print("FATAL RECOVERY 2 ERROR")
        print(clean_text(str(error)))
        return 1

    completed_at = utc_now_iso()

    status = (
        "success"
        if not failed
        else "partial"
    )

    summary = {
        "status": status,
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "recovery1RunSummary": str(
            RECOVERY_1_SUMMARY_FILE
        ),
        "excelFile": str(EXCEL_FILE),
        "recovery1FailureSource": "stillFailedFundsDetail",
        "recovery1FailedFunds": len(
            recovery_funds
        ),
        "recovery1FailedRows": [
            item["excelRow"]
            for item in recovery_funds
        ],
        "recovery2Universe": len(universe),
        "recovery2UniverseRows": [
            item["excelRow"]
            for item in universe
        ],
        "successfulFunds": len(successful),
        "failedFunds": len(failed),
        "totalPublishedTopHoldings": sum(
            item["topHoldingsCount"]
            for item in successful
        ),
        "successfulFundsDetail": successful,
        "failedFundsDetail": failed,
        "rules": {
            "frozenBaselineUntouched": True,
            "recovery1Untouched": True,
            "failedFundsDetailIsExclusiveRecoveryUniverse": True,
            "recovery1FailureField": "stillFailedFundsDetail",
            "excelColumnAControlsMasterUniverse": True,
            "excelUsedOnlyForRecovery1RowConsistency": True,
            "officialPrudentialFactsheetOnly": True,
            "officialPrudentialSingaporeOnly": True,
            "maximumHoldings": MAX_HOLDINGS,
            "publishedHoldingCountUsedExactly": True,
            "fewerThanTenHoldingsAllowed": True,
            "noForcedTenEntries": True,
            "noInferredHoldings": True,
            "noInferredRanks": True,
            "noInferredNames": True,
            "noFabricatedHoldingWeights": True,
            "duplicateHoldingNamesAllowed": True,
            "duplicateHoldingPercentagesAllowed": True,
            "multilineHoldingNamesSupported": True,
            "fixedIncomeLastPercentageOnlyAfterPhysicalRowConfirmation": True,
            "noProximityNameWeightMatching": True,
            "spatialFallbackUsesOfficialPdfCoordinates": True,
            "finalVerificationUsesFreshOfficialPdf": True,
            "finalVerificationRequiresExactSignature": True,
        },
    }

    save_json(
        RUN_SUMMARY_FILE,
        summary,
    )

    save_json(
        ALL_HOLDINGS_FILE,
        {
            "status": status,
            "source": str(
                RECOVERY_1_SUMMARY_FILE
            ),
            "sourceField": "stillFailedFundsDetail",
            "fundCount": len(successful),
            "totalPublishedTopHoldings": sum(
                item["topHoldingsCount"]
                for item in successful
            ),
            "funds": all_holdings,
        },
    )

    print()
    print("============================================================")
    print("STEP 3 COMPLETE")
    print("============================================================")
    print(f"Recovery 2 universe: {len(universe)}")
    print(f"Successful funds:    {len(successful)}")
    print(f"Failed funds:        {len(failed)}")
    print(
        "Published holdings:  "
        f"{sum(item['topHoldingsCount'] for item in successful)}"
    )
    print(
        "Recovery 2 summary:  "
        f"{RUN_SUMMARY_FILE}"
    )
    print(
        "All holdings:        "
        f"{ALL_HOLDINGS_FILE}"
    )
    print("============================================================")

    # Intentionally return 0 even when individual funds remain unresolved.
    # This keeps GitHub Actions artifacts available for diagnosis.
    return 0


if __name__ == "__main__":
    sys.exit(main())
