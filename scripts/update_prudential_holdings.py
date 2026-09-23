#!/usr/bin/env python3

"""
VGrat FMS - Prudential Failed-Fund Top Holdings Recovery

IMPORTANT
=========

This script is a RECOVERY / TEST script.

It MUST NOT modify:

    scripts/test_prudential_holdings.py
    scripts/test_pruaccess.py
    output_holdings/
    data.json

The frozen baseline remains the authoritative result for all funds that
already passed baseline extraction.

This script only attempts funds listed in:

    output_holdings/run_summary.json

under:

    failedFundsDetail

Recovery output is written only to:

    output_holdings_recovery/

PURPOSE
=======

Recover official Prudential Top Holdings for funds that failed the frozen
baseline parser.

The recovery engine deliberately does NOT rely only on PDF reading order.

It uses:

    1. Official Prudential fund URL
    2. Official Prudential fund page
    3. Official factsheet PDF
    4. PDF physical coordinates supplied by pypdf
    5. Visual line reconstruction
    6. Column-aware grouping
    7. Published holding percentage detection
    8. Holding-name reconstruction
    9. Final verification against the original PDF text

NO synthetic data is created.

NO percentages are calculated.

NO holding names are invented.

NO missing percentages are estimated.

NO successful baseline fund is reinterpreted.

The recovery script is intentionally conservative. If it cannot establish
that a name and percentage were actually published together in the official
factsheet, that fund remains failed.

DEPENDENCIES
============

    openpyxl
    pypdf
    playwright

Playwright Chromium must be installed by the GitHub Actions workflow.

Python 3.12 compatible.
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# ============================================================================
# CONFIGURATION
# ============================================================================

BASELINE_SUMMARY = Path("output_holdings/run_summary.json")
EXCEL_FILE = Path("Funds Links.xlsm")

RECOVERY_ROOT = Path("output_holdings_recovery")
RECOVERY_FUNDS = RECOVERY_ROOT / "funds"
RECOVERY_SUMMARY = RECOVERY_ROOT / "run_summary.json"
RECOVERY_ALL_HOLDINGS = RECOVERY_ROOT / "all_holdings.json"

SHEET_NAME = None

PAGE_TIMEOUT_MS = 60_000
FACTSHEET_DOWNLOAD_TIMEOUT = 90

MAX_TOP_HOLDINGS = 10

# A factsheet can contain many unrelated percentages. These thresholds are
# deliberately conservative.
MIN_PERCENT = 0.0
MAX_PERCENT = 100.0

# PDF coordinate tolerances.
Y_TOLERANCE = 3.5
COLUMN_GAP_MIN = 45.0

# Holding names normally have substantially more characters than a bare
# percentage. These are used only to reject obvious non-name fragments.
MIN_HOLDING_NAME_LENGTH = 2

# Weights in factsheets are commonly represented as:
#
#   6.1%
#   6.1 %
#   6.10%
#
# Do not match numbers that are not explicitly followed by %.
PERCENT_RE = re.compile(
    r"(?<![\d.])"
    r"(\d{1,3}(?:\.\d{1,4})?)"
    r"\s*%"
)

RANK_RE = re.compile(
    r"^\s*(?:"
    r"(\d{1,2})[.)]\s+"
    r"|"
    r"(?:#\s*)?(\d{1,2})\s+"
    r")"
)

HOLDINGS_HEADER_PATTERNS = [
    re.compile(r"\btop\s+10\s+holdings\b", re.I),
    re.compile(r"\btop\s+ten\s+holdings\b", re.I),
    re.compile(r"\btop\s+holdings\b", re.I),
    re.compile(r"\btop10\s+holdings\b", re.I),
]

SECTION_STOP_PATTERNS = [
    re.compile(r"^\s*rating\b", re.I),
    re.compile(r"^\s*maturity\b", re.I),
    re.compile(r"^\s*sector\b", re.I),
    re.compile(r"^\s*country\b", re.I),
    re.compile(r"^\s*asset\s+allocation\b", re.I),
    re.compile(r"^\s*geographical\b", re.I),
    re.compile(r"^\s*regional\b", re.I),
    re.compile(r"^\s*portfolio\s+breakdown\b", re.I),
    re.compile(r"^\s*investment\s+breakdown\b", re.I),
    re.compile(r"^\s*credit\s+rating\b", re.I),
    re.compile(r"^\s*duration\b", re.I),
    re.compile(r"^\s*cash\s+and\s+cash\s+equivalents\b", re.I),
]

NON_NAME_PATTERNS = [
    re.compile(r"^\s*top\s+10\s+holdings\b", re.I),
    re.compile(r"^\s*market\s+value\b", re.I),
    re.compile(r"^\s*%?\s*of\s+(?:net\s+)?assets\b", re.I),
    re.compile(r"^\s*weight\b", re.I),
    re.compile(r"^\s*source\s*:", re.I),
    re.compile(r"^\s*inception\s+date\b", re.I),
    re.compile(r"^\s*benchmark\b", re.I),
    re.compile(r"^\s*performance\b", re.I),
]

FACTSHEET_LINK_HINTS = (
    "factsheet",
    "fund-factsheet",
    "fundfactsheet",
    ".pdf",
)


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class PdfToken:
    text: str
    x: float
    y: float
    width: float
    height: float
    page: int


@dataclass
class VisualLine:
    page: int
    y: float
    x_start: float
    x_end: float
    text: str
    tokens: list[PdfToken]


@dataclass
class HoldingCandidate:
    rank: int | None
    name: str
    percentage: float
    page: int
    y: float
    evidence_text: str
    source_line_indexes: list[int]


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(value: str) -> str:
    value = value.replace("\u00a0", " ")
    value = value.replace("\u200b", "")
    value = value.replace("\u00ad", "")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_for_compare(value: str) -> str:
    value = clean_text(value).lower()
    value = value.replace("–", "-")
    value = value.replace("—", "-")
    value = value.replace("’", "'")
    value = value.replace("“", '"')
    value = value.replace("”", '"')
    return value


def safe_filename(value: str, maximum: int = 180) -> str:
    value = clean_text(value)
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", "_", value)
    value = value.strip("._ ")
    if not value:
        value = "unknown"
    return value[:maximum]


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_json(path: Path, data: Any) -> None:
    ensure_dir(path.parent)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def write_text(path: Path, text: str) -> None:
    ensure_dir(path.parent)
    path.write_text(text, encoding="utf-8")


def is_reasonable_percentage(value: float) -> bool:
    return MIN_PERCENT <= value <= MAX_PERCENT


def parse_percentage(text: str) -> float | None:
    match = PERCENT_RE.search(text)
    if not match:
        return None

    try:
        value = float(match.group(1))
    except ValueError:
        return None

    if not is_reasonable_percentage(value):
        return None

    return value


def remove_percentage_tokens(text: str) -> str:
    text = PERCENT_RE.sub("", text)
    return clean_text(text)


def looks_like_non_name(text: str) -> bool:
    text = clean_text(text)

    if not text:
        return True

    for pattern in NON_NAME_PATTERNS:
        if pattern.search(text):
            return True

    if PERCENT_RE.fullmatch(text):
        return True

    if re.fullmatch(r"[\d\s.,%+\-()/]+", text):
        return True

    return False


def is_plausible_holding_name(text: str) -> bool:
    text = clean_text(text)

    if len(text) < MIN_HOLDING_NAME_LENGTH:
        return False

    if looks_like_non_name(text):
        return False

    # Reject strings that are overwhelmingly numeric.
    digits = sum(ch.isdigit() for ch in text)
    letters = sum(ch.isalpha() for ch in text)

    if letters == 0:
        return False

    if digits > letters * 2:
        return False

    return True


def strip_rank_prefix(text: str) -> tuple[int | None, str]:
    match = RANK_RE.match(text)

    if not match:
        return None, text

    rank_text = match.group(1) or match.group(2)

    try:
        rank = int(rank_text)
    except ValueError:
        rank = None

    remainder = text[match.end():].strip()

    return rank, remainder


# ============================================================================
# BASELINE INPUT
# ============================================================================

def load_baseline_summary() -> dict[str, Any]:
    if not BASELINE_SUMMARY.exists():
        raise RuntimeError(
            "Frozen baseline run_summary.json was not found: "
            f"{BASELINE_SUMMARY}"
        )

    with BASELINE_SUMMARY.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise RuntimeError(
            "Frozen baseline run_summary.json is not a JSON object."
        )

    return data


def get_failed_fund_details(summary: dict[str, Any]) -> list[dict[str, Any]]:
    """
    The frozen baseline uses:

        failedFundsDetail: [...]

    This function deliberately requires that exact structure.

    It does NOT silently fall back to:
        failed
        failedFunds

    because those fields have different meanings in the frozen baseline.
    """

    details = summary.get("failedFundsDetail")

    if details is None:
        raise RuntimeError(
            "Frozen baseline does not contain 'failedFundsDetail'. "
            "The recovery script will not guess another structure."
        )

    if not isinstance(details, list):
        raise RuntimeError(
            "'failedFundsDetail' exists but is not a list."
        )

    normalized: list[dict[str, Any]] = []

    for item in details:
        if not isinstance(item, dict):
            raise RuntimeError(
                "An entry in 'failedFundsDetail' is not an object."
            )

        required = (
            "excelRow",
            "prudentialUrl",
            "pruAccessName",
        )

        missing = [key for key in required if key not in item]

        if missing:
            raise RuntimeError(
                "A failed-fund detail entry is missing required fields: "
                + ", ".join(missing)
            )

        normalized.append(item)

    return normalized


# ============================================================================
# EXCEL VERIFICATION
# ============================================================================

def load_excel_funds() -> dict[int, dict[str, Any]]:
    if not EXCEL_FILE.exists():
        raise RuntimeError(f"Excel source was not found: {EXCEL_FILE}")

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_links=True,
    )

    if SHEET_NAME:
        if SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Worksheet '{SHEET_NAME}' was not found in {EXCEL_FILE}"
            )
        worksheet = workbook[SHEET_NAME]
    else:
        worksheet = workbook[workbook.sheetnames[0]]

    funds: dict[int, dict[str, Any]] = {}

    for row_number, row in enumerate(
        worksheet.iter_rows(min_col=1, max_col=2, values_only=True),
        start=1,
    ):
        url = row[0]
        pru_name = row[1] if len(row) > 1 else None

        if row_number == 1:
            continue

        if url is None:
            continue

        url_text = str(url).strip()

        if not url_text:
            continue

        funds[row_number] = {
            "excelRow": row_number,
            "prudentialUrl": url_text,
            "pruAccessName": clean_text(str(pru_name or "")),
        }

    workbook.close()

    return funds


def verify_failure_against_excel(
    failure: dict[str, Any],
    excel_funds: dict[int, dict[str, Any]],
) -> dict[str, Any]:
    row = int(failure["excelRow"])

    excel_record = excel_funds.get(row)

    if not excel_record:
        raise RuntimeError(
            f"Excel row {row} from failedFundsDetail does not exist "
            "in Funds Links.xlsm."
        )

    baseline_url = clean_text(str(failure["prudentialUrl"]))
    excel_url = clean_text(str(excel_record["prudentialUrl"]))

    if baseline_url != excel_url:
        raise RuntimeError(
            f"Excel URL mismatch for row {row}.\n"
            f"Baseline: {baseline_url}\n"
            f"Excel:    {excel_url}"
        )

    baseline_name = clean_text(str(failure["pruAccessName"] or ""))
    excel_name = clean_text(str(excel_record["pruAccessName"] or ""))

    if baseline_name and excel_name and baseline_name != excel_name:
        raise RuntimeError(
            f"Excel PruAccess-name mismatch for row {row}.\n"
            f"Baseline: {baseline_name}\n"
            f"Excel:    {excel_name}"
        )

    return excel_record


# ============================================================================
# PRUDENTIAL PAGE / FACTSHEET
# ============================================================================

def normalize_url(url: str) -> str:
    parsed = urlparse(url)

    if not parsed.scheme:
        return "https://" + url

    return url


def fetch_bytes(url: str) -> bytes:
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/153.0 Safari/537.36"
            ),
            "Accept": "application/pdf,*/*",
        },
    )

    with urlopen(request, timeout=FACTSHEET_DOWNLOAD_TIMEOUT) as response:
        return response.read()


def looks_like_pdf(data: bytes) -> bool:
    return data[:5] == b"%PDF-"


def extract_factsheet_candidates_from_page(
    page,
    base_url: str,
) -> list[str]:
    hrefs: list[str] = []

    try:
        links = page.locator("a").evaluate_all(
            """
            elements => elements.map(a => ({
                href: a.href || "",
                text: (a.innerText || "").trim(),
                aria: a.getAttribute("aria-label") || "",
                title: a.getAttribute("title") || ""
            }))
            """
        )
    except Exception:
        links = []

    for link in links:
        href = str(link.get("href") or "").strip()

        if not href:
            continue

        combined = " ".join(
            [
                href,
                str(link.get("text") or ""),
                str(link.get("aria") or ""),
                str(link.get("title") or ""),
            ]
        ).lower()

        if (
            ".pdf" in href.lower()
            or "factsheet" in combined
            or "fund factsheet" in combined
        ):
            absolute = urljoin(base_url, href)

            if absolute not in hrefs:
                hrefs.append(absolute)

    return hrefs


def download_official_factsheet(
    page,
    prudential_url: str,
    output_pdf: Path,
) -> tuple[str, bytes]:
    """
    Locate the official Prudential factsheet from the official fund page.

    First attempts to capture a PDF download through the browser.

    Then falls back to direct download of candidate official PDF links.

    Only PDF content is accepted.
    """

    candidates = extract_factsheet_candidates_from_page(
        page,
        prudential_url,
    )

    # Put direct .pdf links first.
    candidates.sort(
        key=lambda value: (
            0 if ".pdf" in value.lower() else 1,
            0 if "factsheet" in value.lower() else 1,
        )
    )

    # Browser-triggered downloads can be necessary when hrefs are generated.
    try:
        fact_links = page.get_by_text(
            re.compile(r"factsheet", re.I)
        )

        count = fact_links.count()

        for index in range(min(count, 10)):
            locator = fact_links.nth(index)

            try:
                with page.expect_download(
                    timeout=7_000
                ) as download_info:
                    locator.click()

                download = download_info.value

                temp_path = Path(
                    download.path() or ""
                )

                if temp_path.exists():
                    data = temp_path.read_bytes()

                    if looks_like_pdf(data):
                        ensure_dir(output_pdf.parent)
                        output_pdf.write_bytes(data)

                        return (
                            download.suggested_filename
                            or "factsheet.pdf",
                            data,
                        )

            except Exception:
                continue

    except Exception:
        pass

    # Direct candidate download.
    for candidate in candidates:
        try:
            data = fetch_bytes(candidate)

            if not looks_like_pdf(data):
                continue

            ensure_dir(output_pdf.parent)
            output_pdf.write_bytes(data)

            return candidate, data

        except Exception:
            continue

    # Last resort: search HTML source for PDF-like URLs.
    try:
        html = page.content()

        raw_pdf_urls = re.findall(
            r"""["']([^"']+\.pdf(?:\?[^"']*)?)["']""",
            html,
            flags=re.I,
        )

        for raw in raw_pdf_urls:
            candidate = urljoin(prudential_url, raw)

            try:
                data = fetch_bytes(candidate)

                if not looks_like_pdf(data):
                    continue

                ensure_dir(output_pdf.parent)
                output_pdf.write_bytes(data)

                return candidate, data

            except Exception:
                continue

    except Exception:
        pass

    raise RuntimeError(
        "Could not locate or download an official Prudential factsheet PDF."
    )


# ============================================================================
# PDF TEXT / COORDINATES
# ============================================================================

def extract_pdf_tokens(pdf_path: Path) -> list[PdfToken]:
    reader = PdfReader(str(pdf_path))

    tokens: list[PdfToken] = []

    for page_number, page in enumerate(reader.pages, start=1):
        current_text_parts: list[str] = []
        current_x: float | None = None
        current_y: float | None = None
        current_width: float = 0.0
        current_height: float = 0.0

        def visitor_text(
            text: str,
            cm,
            tm,
            font_dict,
            font_size,
        ):
            nonlocal current_text_parts
            nonlocal current_x
            nonlocal current_y
            nonlocal current_width
            nonlocal current_height

            if not text:
                return

            x = float(tm[4])
            y = float(tm[5])

            # pypdf can emit a text chunk containing multiple characters,
            # spaces and line breaks. We preserve it as one physical token.
            cleaned = text.replace("\r", "\n")

            pieces = cleaned.split("\n")

            for piece_index, piece in enumerate(pieces):
                piece = piece.strip()

                if not piece:
                    continue

                estimated_width = max(
                    float(font_size) * 0.45 * len(piece),
                    1.0,
                )

                tokens.append(
                    PdfToken(
                        text=piece,
                        x=x,
                        y=y - (piece_index * float(font_size)),
                        width=estimated_width,
                        height=float(font_size),
                        page=page_number,
                    )
                )

            current_text_parts = []
            current_x = None
            current_y = None
            current_width = 0.0
            current_height = 0.0

        try:
            page.extract_text(
                visitor_text=visitor_text
            )
        except Exception:
            # Some malformed PDFs may not support visitor extraction.
            # The caller will handle the empty-token condition.
            continue

    return tokens


def extract_plain_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))

    pages: list[str] = []

    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")

    return "\n".join(pages)


def build_visual_lines(tokens: list[PdfToken]) -> list[VisualLine]:
    grouped: dict[tuple[int, int], list[PdfToken]] = defaultdict(list)

    for token in tokens:
        # Quantise Y only for grouping. Original Y is preserved later.
        y_bucket = round(token.y / Y_TOLERANCE)

        grouped[(token.page, y_bucket)].append(token)

    lines: list[VisualLine] = []

    for (page, _bucket), group in grouped.items():
        group.sort(key=lambda token: token.x)

        # The bucket can still occasionally contain two physical rows.
        # Split when the Y difference is clearly larger than tolerance.
        subgroups: list[list[PdfToken]] = []

        for token in group:
            placed = False

            for subgroup in subgroups:
                avg_y = sum(t.y for t in subgroup) / len(subgroup)

                if abs(token.y - avg_y) <= Y_TOLERANCE:
                    subgroup.append(token)
                    placed = True
                    break

            if not placed:
                subgroups.append([token])

        for subgroup in subgroups:
            subgroup.sort(key=lambda token: token.x)

            text = clean_text(
                " ".join(token.text for token in subgroup)
            )

            if not text:
                continue

            x_start = min(token.x for token in subgroup)

            x_end = max(
                token.x + token.width
                for token in subgroup
            )

            y = sum(token.y for token in subgroup) / len(subgroup)

            lines.append(
                VisualLine(
                    page=page,
                    y=y,
                    x_start=x_start,
                    x_end=x_end,
                    text=text,
                    tokens=subgroup,
                )
            )

    lines.sort(
        key=lambda line: (
            line.page,
            -line.y,
            line.x_start,
        )
    )

    return lines


# ============================================================================
# HOLDINGS SECTION LOCATION
# ============================================================================

def find_holdings_header_indexes(
    lines: list[VisualLine],
) -> list[int]:
    indexes: list[int] = []

    for index, line in enumerate(lines):
        text = clean_text(line.text)

        for pattern in HOLDINGS_HEADER_PATTERNS:
            if pattern.search(text):
                indexes.append(index)
                break

    return indexes


def is_section_stop(text: str) -> bool:
    text = clean_text(text)

    for pattern in SECTION_STOP_PATTERNS:
        if pattern.search(text):
            return True

    return False


def choose_best_holdings_section(
    lines: list[VisualLine],
) -> tuple[int, int]:
    """
    Returns [start_index, end_index).

    The parser considers every Top Holdings header and chooses the section
    containing the strongest sequence of explicit percentages and names.

    This is important for the bond funds where the PDF reading order can
    interleave another portfolio table with the Top Holdings table.
    """

    headers = find_holdings_header_indexes(lines)

    if not headers:
        raise RuntimeError(
            "Could not locate a Top Holdings section in the official PDF."
        )

    sections: list[tuple[int, int, int]] = []

    for position, start in enumerate(headers):
        next_header = (
            headers[position + 1]
            if position + 1 < len(headers)
            else len(lines)
        )

        stop = next_header

        for index in range(start + 1, next_header):
            if is_section_stop(lines[index].text):
                stop = index
                break

        candidate_lines = lines[start + 1:stop]

        explicit_percentages = sum(
            1
            for line in candidate_lines
            if parse_percentage(line.text) is not None
        )

        nonempty_names = sum(
            1
            for line in candidate_lines
            if is_plausible_holding_name(
                remove_percentage_tokens(line.text)
            )
        )

        score = (
            explicit_percentages * 5
            + min(nonempty_names, 20)
        )

        sections.append((score, start, stop))

    sections.sort(
        key=lambda item: (
            item[0],
            -item[1],
        ),
        reverse=True,
    )

    _, start, stop = sections[0]

    return start, stop


# ============================================================================
# COLUMN DETECTION
# ============================================================================

def estimate_columns(
    lines: list[VisualLine],
) -> list[tuple[float, float]]:
    """
    Estimate visual columns from the X coordinates.

    The purpose is not to perfectly reconstruct the PDF page.

    It is to prevent a two-column factsheet from combining the left-column
    holding name with the right-column percentage belonging to another row.
    """

    if not lines:
        return []

    starts = sorted(
        line.x_start
        for line in lines
        if line.text
    )

    clusters: list[list[float]] = []

    for x in starts:
        placed = False

        for cluster in clusters:
            center = sum(cluster) / len(cluster)

            if abs(x - center) <= COLUMN_GAP_MIN:
                cluster.append(x)
                placed = True
                break

        if not placed:
            clusters.append([x])

    # Collapse clusters that are too close.
    changed = True

    while changed:
        changed = False

        if len(clusters) <= 1:
            break

        clusters.sort(
            key=lambda cluster: sum(cluster) / len(cluster)
        )

        merged: list[list[float]] = []

        for cluster in clusters:
            if not merged:
                merged.append(cluster)
                continue

            previous = merged[-1]

            prev_center = sum(previous) / len(previous)
            curr_center = sum(cluster) / len(cluster)

            if curr_center - prev_center < COLUMN_GAP_MIN:
                merged[-1].extend(cluster)
                changed = True
            else:
                merged.append(cluster)

        clusters = merged

    centers = sorted(
        sum(cluster) / len(cluster)
        for cluster in clusters
    )

    if len(centers) > 3:
        # Most factsheets are either single-column or two-column for this
        # section. More clusters are usually artifacts of text alignment.
        centers = centers[:3]

    result: list[tuple[float, float]] = []

    if len(centers) == 1:
        return [(centers[0] - 500, centers[0] + 500)]

    for index, center in enumerate(centers):
        left = (
            -10_000
            if index == 0
            else (centers[index - 1] + center) / 2
        )

        right = (
            10_000
            if index == len(centers) - 1
            else (center + centers[index + 1]) / 2
        )

        result.append((left, right))

    return result


def assign_column(
    line: VisualLine,
    columns: list[tuple[float, float]],
) -> int:
    if not columns:
        return 0

    center = (line.x_start + line.x_end) / 2

    distances = []

    for index, (left, right) in enumerate(columns):
        if left <= center <= right:
            return index

        if center < left:
            distance = left - center
        else:
            distance = center - right

        distances.append((distance, index))

    distances.sort()

    return distances[0][1]


# ============================================================================
# VISUAL HOLDING RECONSTRUCTION
# ============================================================================

def split_percentage_from_line(
    text: str,
) -> tuple[str, float | None]:
    """
    Remove the LAST explicit percentage from a visual line.

    This is intentional.

    Example:

        SECURITY 3.25% 01/06/2031 4.52%

    becomes:

        SECURITY 3.25% 01/06/2031
        4.52

    This preserves coupon/rate percentages inside fixed-income names.
    """

    matches = list(PERCENT_RE.finditer(text))

    if not matches:
        return clean_text(text), None

    match = matches[-1]

    try:
        percentage = float(match.group(1))
    except ValueError:
        return clean_text(text), None

    if not is_reasonable_percentage(percentage):
        return clean_text(text), None

    name_part = (
        text[:match.start()]
        + " "
        + text[match.end():]
    )

    return clean_text(name_part), percentage


def line_contains_rank(line: VisualLine) -> tuple[int | None, str]:
    return strip_rank_prefix(line.text)


def reconstruct_candidates(
    section_lines: list[VisualLine],
) -> list[HoldingCandidate]:
    """
    Reconstruct published holdings from the physical PDF layout.

    Strategy:

    A. Keep physical page/column order.
    B. Identify explicit percentages.
    C. If the same line contains a name, pair immediately.
    D. If a percentage is detached onto another line, associate it with
       the nearest preceding name in the same visual column.
    E. Preserve multi-line names.
    F. Never invent a percentage.
    G. Never calculate a percentage.
    """

    if not section_lines:
        return []

    columns = estimate_columns(section_lines)

    column_lines: dict[int, list[tuple[int, VisualLine]]] = defaultdict(list)

    for original_index, line in enumerate(section_lines):
        column = assign_column(line, columns)
        column_lines[column].append(
            (original_index, line)
        )

    candidates: list[HoldingCandidate] = []

    for column_index in sorted(column_lines):
        items = column_lines[column_index]

        # PDF coordinate order is top to bottom.
        items.sort(
            key=lambda item: (
                -item[1].y,
                item[1].x_start,
            )
        )

        pending_name_lines: list[tuple[int, VisualLine]] = []
        pending_rank: int | None = None

        for local_index, (original_index, line) in enumerate(items):
            raw = clean_text(line.text)

            if not raw:
                continue

            rank, rankless = line_contains_rank(line)

            if rank is not None:
                # A new rank means a new logical holding.
                if pending_name_lines:
                    # Do not throw away a possible detached percentage yet.
                    # It is handled by look-ahead below.
                    pass

                pending_rank = rank
                raw = clean_text(rankless)

            name_part, percentage = split_percentage_from_line(raw)

            # ----------------------------------------------------------------
            # CASE 1:
            # Same visual line contains both name and percentage.
            # ----------------------------------------------------------------
            if percentage is not None and is_plausible_holding_name(name_part):
                candidate = HoldingCandidate(
                    rank=pending_rank,
                    name=name_part,
                    percentage=percentage,
                    page=line.page,
                    y=line.y,
                    evidence_text=line.text,
                    source_line_indexes=[original_index],
                )

                candidates.append(candidate)

                pending_name_lines = []
                pending_rank = None
                continue

            # ----------------------------------------------------------------
            # CASE 2:
            # This line contains only a percentage.
            #
            # Associate it with the immediately preceding logical name.
            # ----------------------------------------------------------------
            if percentage is not None and not is_plausible_holding_name(
                name_part
            ):
                if pending_name_lines:
                    name = clean_text(
                        " ".join(
                            item[1].text
                            for item in pending_name_lines
                        )
                    )

                    if is_plausible_holding_name(name):
                        first_line = pending_name_lines[0][1]

                        candidates.append(
                            HoldingCandidate(
                                rank=pending_rank,
                                name=name,
                                percentage=percentage,
                                page=line.page,
                                y=line.y,
                                evidence_text=(
                                    name
                                    + " | "
                                    + line.text
                                ),
                                source_line_indexes=[
                                    item[0]
                                    for item in pending_name_lines
                                ] + [original_index],
                            )
                        )

                        pending_name_lines = []
                        pending_rank = None
                        continue

                # No valid preceding name.
                # Preserve diagnostic information by simply not creating
                # a fabricated candidate.
                continue

            # ----------------------------------------------------------------
            # CASE 3:
            # Name-only line.
            #
            # Keep it pending for:
            #
            #   - a continuation line
            #   - a detached percentage
            # ----------------------------------------------------------------
            if is_plausible_holding_name(name_part):
                # Avoid swallowing obvious section labels.
                if is_section_stop(name_part):
                    pending_name_lines = []
                    pending_rank = None
                    continue

                pending_name_lines.append(
                    (original_index, line)
                )

                # Limit runaway capture.
                if len(pending_name_lines) > 5:
                    pending_name_lines = pending_name_lines[-5:]

    # ------------------------------------------------------------------------
    # Remove obvious duplicates caused by PDF coordinate fragments.
    # Preserve legitimate duplicate holdings when the evidence differs.
    # ------------------------------------------------------------------------

    deduped: list[HoldingCandidate] = []

    seen_exact: set[tuple[str, float, int]] = set()

    for candidate in candidates:
        key = (
            normalize_for_compare(candidate.name),
            round(candidate.percentage, 8),
            candidate.page,
        )

        if key in seen_exact:
            continue

        seen_exact.add(key)
        deduped.append(candidate)

    # ------------------------------------------------------------------------
    # Rank recovery:
    #
    # Published factsheets may not explicitly print 1-10.
    # In that case rank is the published visual order.
    # ------------------------------------------------------------------------

    if deduped:
        for index, candidate in enumerate(deduped, start=1):
            if candidate.rank is None:
                candidate.rank = index

    return deduped


# ============================================================================
# SECONDARY RECONSTRUCTION STRATEGIES
# ============================================================================

def reconstruct_from_plain_lines(
    section_lines: list[VisualLine],
) -> list[HoldingCandidate]:
    """
    Conservative fallback.

    This is not the frozen baseline parser.

    It is used only after coordinate reconstruction.

    It handles PDFs where pypdf's visitor coordinates are poor but normal
    text extraction still preserves useful line relationships.
    """

    candidates: list[HoldingCandidate] = []

    pending_name_lines: list[tuple[int, VisualLine]] = []
    pending_rank: int | None = None

    for index, line in enumerate(section_lines):
        raw = clean_text(line.text)

        if not raw:
            continue

        rank, raw_without_rank = strip_rank_prefix(raw)

        if rank is not None:
            pending_rank = rank
            raw = raw_without_rank

        name_part, percentage = split_percentage_from_line(raw)

        if percentage is not None:
            if is_plausible_holding_name(name_part):
                candidates.append(
                    HoldingCandidate(
                        rank=pending_rank,
                        name=name_part,
                        percentage=percentage,
                        page=line.page,
                        y=line.y,
                        evidence_text=line.text,
                        source_line_indexes=[index],
                    )
                )

                pending_name_lines = []
                pending_rank = None
                continue

            if pending_name_lines:
                name = clean_text(
                    " ".join(
                        item[1].text
                        for item in pending_name_lines
                    )
                )

                if is_plausible_holding_name(name):
                    candidates.append(
                        HoldingCandidate(
                            rank=pending_rank,
                            name=name,
                            percentage=percentage,
                            page=line.page,
                            y=line.y,
                            evidence_text=(
                                name
                                + " | "
                                + line.text
                            ),
                            source_line_indexes=[
                                item[0]
                                for item in pending_name_lines
                            ] + [index],
                        )
                    )

                    pending_name_lines = []
                    pending_rank = None
                    continue

        if is_plausible_holding_name(name_part):
            pending_name_lines.append((index, line))

            if len(pending_name_lines) > 5:
                pending_name_lines = pending_name_lines[-5:]

    for index, candidate in enumerate(candidates, start=1):
        if candidate.rank is None:
            candidate.rank = index

    return candidates


def validate_candidate_sequence(
    candidates: list[HoldingCandidate],
) -> tuple[bool, str]:
    if not candidates:
        return False, "No holding candidates were reconstructed."

    if len(candidates) > MAX_TOP_HOLDINGS:
        # A section may contain unrelated percentages after the holdings.
        # Only accept the first 10 if they form a coherent sequence.
        candidates = candidates[:MAX_TOP_HOLDINGS]

    for candidate in candidates:
        if not is_plausible_holding_name(candidate.name):
            return (
                False,
                "A reconstructed holding name failed validation."
            )

        if not is_reasonable_percentage(candidate.percentage):
            return (
                False,
                "A reconstructed holding percentage failed validation."
            )

    return True, ""


# ============================================================================
# PDF VERIFICATION
# ============================================================================

def verify_candidate_against_full_pdf(
    candidate: HoldingCandidate,
    full_pdf_text: str,
) -> tuple[bool, str]:
    """
    Final anti-fabrication verification.

    Both the name and exact percentage must be independently recoverable
    from the official PDF text.

    The script does not accept a candidate solely because it came from
    coordinates.
    """

    normalized_pdf = normalize_for_compare(full_pdf_text)

    normalized_name = normalize_for_compare(candidate.name)

    if normalized_name not in normalized_pdf:
        # Hyphenation can split words in PDF text.
        compact_name = normalized_name.replace("-", " ")
        compact_pdf = normalized_pdf.replace("-", " ")

        if compact_name not in compact_pdf:
            return (
                False,
                "Holding name was not found in the full official PDF text."
            )

    percentage_text = (
        f"{candidate.percentage:g}%"
    )

    if normalize_for_compare(percentage_text) not in normalized_pdf:
        return (
            False,
            "Published holding percentage was not found in the "
            "full official PDF text."
        )

    return True, ""


def verify_name_percentage_proximity(
    candidate: HoldingCandidate,
    full_pdf_text: str,
) -> bool:
    """
    Secondary verification.

    The PDF may use line breaks between the security name and percentage.

    We therefore search for the name followed within a bounded amount of
    text by the exact percentage, and also the reverse order.

    This is a verification aid only. It never creates data.
    """

    name = normalize_for_compare(candidate.name)
    percentage = normalize_for_compare(
        f"{candidate.percentage:g}%"
    )

    text = normalize_for_compare(full_pdf_text)

    # Collapse common whitespace.
    text = re.sub(r"\s+", " ", text)
    name = re.sub(r"\s+", " ", name)

    # Allow up to 160 characters between name and percentage.
    forward = re.search(
        re.escape(name) + r".{0,160}" + re.escape(percentage),
        text,
    )

    if forward:
        return True

    reverse = re.search(
        re.escape(percentage) + r".{0,160}" + re.escape(name),
        text,
    )

    return bool(reverse)


# ============================================================================
# HOLDING OUTPUT
# ============================================================================

def candidate_to_output(
    candidate: HoldingCandidate,
    verification_method: str,
) -> dict[str, Any]:
    return {
        "rank": candidate.rank,
        "name": candidate.name,
        "percentage": candidate.percentage,
        "page": candidate.page,
        "pdfY": round(candidate.y, 3),
        "evidenceText": candidate.evidence_text,
        "verification": verification_method,
        "sourceLineIndexes": candidate.source_line_indexes,
    }


def select_final_holdings(
    candidates: list[HoldingCandidate],
    full_pdf_text: str,
) -> tuple[list[dict[str, Any]], list[str]]:
    diagnostics: list[str] = []
    verified: list[HoldingCandidate] = []

    for candidate in candidates:
        ok, reason = verify_candidate_against_full_pdf(
            candidate,
            full_pdf_text,
        )

        if not ok:
            diagnostics.append(
                f"Rejected candidate '{candidate.name}': {reason}"
            )
            continue

        if not verify_name_percentage_proximity(
            candidate,
            full_pdf_text,
        ):
            diagnostics.append(
                f"Rejected candidate '{candidate.name}': "
                "name/percentage proximity could not be verified."
            )
            continue

        verified.append(candidate)

    if not verified:
        return [], diagnostics

    # Never silently accept more than ten.
    verified = verified[:MAX_TOP_HOLDINGS]

    # Published order is preserved.
    output = [
        candidate_to_output(
            candidate,
            "official_pdf_text_and_proximity",
        )
        for candidate in verified
    ]

    return output, diagnostics


# ============================================================================
# DIAGNOSTIC FILES
# ============================================================================

def write_pdf_diagnostics(
    fund_dir: Path,
    pdf_path: Path,
    tokens: list[PdfToken],
    lines: list[VisualLine],
    section_start: int | None,
    section_end: int | None,
) -> None:
    token_data = [
        {
            "page": token.page,
            "x": round(token.x, 3),
            "y": round(token.y, 3),
            "width": round(token.width, 3),
            "height": round(token.height, 3),
            "text": token.text,
        }
        for token in tokens
    ]

    write_json(
        fund_dir / "pdf_tokens.json",
        token_data,
    )

    line_data = []

    for index, line in enumerate(lines):
        line_data.append(
            {
                "index": index,
                "page": line.page,
                "y": round(line.y, 3),
                "xStart": round(line.x_start, 3),
                "xEnd": round(line.x_end, 3),
                "text": line.text,
                "inHoldingsSection": (
                    section_start is not None
                    and section_end is not None
                    and section_start <= index < section_end
                ),
            }
        )

    write_json(
        fund_dir / "pdf_visual_lines.json",
        line_data,
    )


# ============================================================================
# FUND RECOVERY
# ============================================================================

def recover_single_fund(
    failure: dict[str, Any],
    excel_record: dict[str, Any],
    browser,
) -> dict[str, Any]:
    row = int(failure["excelRow"])

    prudential_url = normalize_url(
        str(failure["prudentialUrl"])
    )

    pru_access_name = clean_text(
        str(failure.get("pruAccessName") or "")
    )

    identifier = safe_filename(
        pru_access_name or f"row_{row}"
    )

    fund_dir = RECOVERY_FUNDS / f"{row}_{identifier}"

    ensure_dir(fund_dir)

    started_at = utc_now()

    page = browser.new_page()

    page.set_default_timeout(PAGE_TIMEOUT_MS)

    metadata: dict[str, Any] = {
        "excelRow": row,
        "prudentialUrl": prudential_url,
        "pruAccessName": pru_access_name,
        "excelRecord": excel_record,
        "baselineError": failure.get("error"),
        "baselineOutputDirectory": failure.get("outputDirectory"),
        "recoveryStartedAtUtc": started_at,
    }

    try:
        # --------------------------------------------------------------------
        # 1. Open official Prudential page.
        # --------------------------------------------------------------------
        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        try:
            page.wait_for_load_state(
                "networkidle",
                timeout=15_000,
            )
        except PlaywrightTimeoutError:
            pass

        time.sleep(1.0)

        metadata["resolvedUrl"] = page.url

        write_text(
            fund_dir / "fund_page.html",
            page.content(),
        )

        # --------------------------------------------------------------------
        # 2. Locate and download official factsheet.
        # --------------------------------------------------------------------
        factsheet_path = fund_dir / "factsheet.pdf"

        factsheet_source, pdf_bytes = download_official_factsheet(
            page,
            prudential_url,
            factsheet_path,
        )

        metadata["factsheetSource"] = factsheet_source
        metadata["factsheetBytes"] = len(pdf_bytes)
        metadata["factsheetSha256"] = sha256_bytes(pdf_bytes)

        # --------------------------------------------------------------------
        # 3. Read official PDF.
        # --------------------------------------------------------------------
        full_pdf_text = extract_plain_pdf_text(
            factsheet_path
        )

        write_text(
            fund_dir / "factsheet_text.txt",
            full_pdf_text,
        )

        if not full_pdf_text.strip():
            raise RuntimeError(
                "Official factsheet PDF contained no extractable text."
            )

        # --------------------------------------------------------------------
        # 4. Extract physical coordinates.
        # --------------------------------------------------------------------
        tokens = extract_pdf_tokens(
            factsheet_path
        )

        if not tokens:
            raise RuntimeError(
                "Could not extract physical PDF text coordinates."
            )

        visual_lines = build_visual_lines(tokens)

        if not visual_lines:
            raise RuntimeError(
                "Could not reconstruct visual PDF lines."
            )

        # --------------------------------------------------------------------
        # 5. Locate best Top Holdings section.
        # --------------------------------------------------------------------
        section_start, section_end = choose_best_holdings_section(
            visual_lines
        )

        section_lines = visual_lines[
            section_start:section_end
        ]

        section_text = "\n".join(
            line.text
            for line in section_lines
        )

        write_text(
            fund_dir / "top_holdings_section.txt",
            section_text,
        )

        write_pdf_diagnostics(
            fund_dir,
            factsheet_path,
            tokens,
            visual_lines,
            section_start,
            section_end,
        )

        # --------------------------------------------------------------------
        # 6. Primary coordinate-aware reconstruction.
        # --------------------------------------------------------------------
        candidates = reconstruct_candidates(
            section_lines
        )

        primary_ok, primary_reason = validate_candidate_sequence(
            candidates
        )

        # --------------------------------------------------------------------
        # 7. Secondary conservative reconstruction if necessary.
        # --------------------------------------------------------------------
        parser_used = "coordinate_visual_reconstruction"

        if not primary_ok or not candidates:
            fallback_candidates = reconstruct_from_plain_lines(
                section_lines
            )

            fallback_ok, fallback_reason = validate_candidate_sequence(
                fallback_candidates
            )

            if fallback_ok:
                candidates = fallback_candidates
                parser_used = "coordinate_section_plain_line_fallback"
            else:
                raise RuntimeError(
                    "Recovery parser could not reconstruct a valid holding "
                    "sequence.\n"
                    f"Primary: {primary_reason}\n"
                    f"Fallback: {fallback_reason}"
                )

        # --------------------------------------------------------------------
        # 8. Final official-PDF verification.
        # --------------------------------------------------------------------
        final_holdings, diagnostics = select_final_holdings(
            candidates,
            full_pdf_text,
        )

        write_json(
            fund_dir / "candidate_holdings.json",
            [
                candidate_to_output(
                    candidate,
                    "candidate_only",
                )
                for candidate in candidates
            ],
        )

        write_json(
            fund_dir / "verification_diagnostics.json",
            diagnostics,
        )

        if not final_holdings:
            raise RuntimeError(
                "No reconstructed holding passed final official-PDF "
                "verification."
            )

        # --------------------------------------------------------------------
        # 9. Strict final checks.
        # --------------------------------------------------------------------
        if len(final_holdings) > MAX_TOP_HOLDINGS:
            raise RuntimeError(
                "Recovery produced more than 10 holdings."
            )

        for holding in final_holdings:
            if not holding["name"]:
                raise RuntimeError(
                    "Recovery produced an empty holding name."
                )

            if holding["percentage"] is None:
                raise RuntimeError(
                    "Recovery produced a holding without a published "
                    "percentage."
                )

        # --------------------------------------------------------------------
        # 10. Save verified holdings.
        # --------------------------------------------------------------------
        write_json(
            fund_dir / "top_holdings.json",
            {
                "status": "success",
                "excelRow": row,
                "prudentialUrl": prudential_url,
                "pruAccessName": pru_access_name,
                "parser": parser_used,
                "verificationPass": True,
                "holdingsCount": len(final_holdings),
                "holdings": final_holdings,
            },
        )

        metadata.update(
            {
                "status": "success",
                "parser": parser_used,
                "verificationPass": True,
                "topHoldingsCount": len(final_holdings),
                "recoveryCompletedAtUtc": utc_now(),
            }
        )

        write_json(
            fund_dir / "metadata.json",
            metadata,
        )

        return {
            "status": "success",
            "excelRow": row,
            "prudentialUrl": prudential_url,
            "fundName": pru_access_name,
            "pruAccessName": pru_access_name,
            "topHoldingsCount": len(final_holdings),
            "holdings": final_holdings,
            "parser": parser_used,
            "verificationPass": True,
            "outputDirectory": str(fund_dir),
        }

    except Exception as exc:
        error_text = str(exc)

        metadata.update(
            {
                "status": "failed",
                "verificationPass": False,
                "error": error_text,
                "recoveryCompletedAtUtc": utc_now(),
            }
        )

        write_json(
            fund_dir / "metadata.json",
            metadata,
        )

        return {
            "status": "failed",
            "excelRow": row,
            "prudentialUrl": prudential_url,
            "fundName": pru_access_name,
            "pruAccessName": pru_access_name,
            "error": error_text,
            "verificationPass": False,
            "outputDirectory": str(fund_dir),
        }

    finally:
        try:
            page.close()
        except Exception:
            pass


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:
    print()
    print("=" * 72)
    print("VGrat FMS - PRUDENTIAL FAILED-FUND HOLDINGS RECOVERY")
    print("=" * 72)
    print()
    print("IMPORTANT:")
    print("  Frozen baseline will NOT be modified.")
    print("  PruAccess script will NOT be modified.")
    print("  Baseline output will NOT be overwritten.")
    print()
    print("Recovery output:")
    print(f"  {RECOVERY_ROOT}")
    print()

    ensure_dir(RECOVERY_ROOT)
    ensure_dir(RECOVERY_FUNDS)

    # ------------------------------------------------------------------------
    # Read exact frozen baseline structure.
    # ------------------------------------------------------------------------
    print("=" * 72)
    print("READING FROZEN BASELINE")
    print("=" * 72)

    baseline = load_baseline_summary()

    print(
        f"Baseline status: "
        f"{baseline.get('status')}"
    )

    print(
        f"Baseline successful funds: "
        f"{baseline.get('successfulFunds')}"
    )

    print(
        f"Baseline failedFunds count: "
        f"{baseline.get('failedFunds')}"
    )

    failed_details = get_failed_fund_details(
        baseline
    )

    print(
        f"Baseline failedFundsDetail entries: "
        f"{len(failed_details)}"
    )

    if not failed_details:
        summary = {
            "status": "success",
            "generatedAtUtc": utc_now(),
            "message": "Frozen baseline contains no failed funds.",
            "failedFundsFromBaseline": 0,
            "matchedRecoveryTargets": 0,
            "successfulFunds": 0,
            "failedFunds": 0,
            "recovered": [],
            "failed": [],
        }

        write_json(
            RECOVERY_SUMMARY,
            summary,
        )

        write_json(
            RECOVERY_ALL_HOLDINGS,
            [],
        )

        print()
        print("No recovery targets.")
        return 0

    # ------------------------------------------------------------------------
    # Read Excel universe.
    # ------------------------------------------------------------------------
    print()
    print("=" * 72)
    print("READING FUNDS LINKS.XLSM")
    print("=" * 72)

    excel_funds = load_excel_funds()

    print(
        f"Excel populated fund rows: "
        f"{len(excel_funds)}"
    )

    # ------------------------------------------------------------------------
    # Validate every baseline failure against Excel.
    # ------------------------------------------------------------------------
    recovery_targets: list[tuple[dict[str, Any], dict[str, Any]]] = []

    for failure in failed_details:
        excel_record = verify_failure_against_excel(
            failure,
            excel_funds,
        )

        recovery_targets.append(
            (failure, excel_record)
        )

    print(
        f"Validated recovery targets: "
        f"{len(recovery_targets)}"
    )

    # ------------------------------------------------------------------------
    # Start browser.
    # ------------------------------------------------------------------------
    recovered: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []

    print()
    print("=" * 72)
    print("STARTING RECOVERY")
    print("=" * 72)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=True,
        )

        try:
            total = len(recovery_targets)

            for index, (failure, excel_record) in enumerate(
                recovery_targets,
                start=1,
            ):
                row = failure["excelRow"]
                name = failure["pruAccessName"]

                print()
                print("-" * 72)
                print(
                    f"[{index}/{total}] "
                    f"Recovering Excel row {row}"
                )
                print(
                    f"Fund: {name}"
                )
                print(
                    f"Baseline error: "
                    f"{failure.get('error', '')}"
                )
                print("-" * 72)

                result = recover_single_fund(
                    failure,
                    excel_record,
                    browser,
                )

                if result["status"] == "success":
                    recovered.append(result)

                    print()
                    print(
                        f"SUCCESS - Row {row}: "
                        f"{result['topHoldingsCount']} holdings"
                    )

                    for holding in result["holdings"]:
                        print(
                            f"  {holding['rank']}. "
                            f"{holding['name']} "
                            f"— {holding['percentage']}%"
                        )

                else:
                    failed.append(result)

                    print()
                    print(
                        f"FAILED - Row {row}: "
                        f"{result['error']}"
                    )

        finally:
            browser.close()

    # ------------------------------------------------------------------------
    # Combined verified recovery holdings.
    # ------------------------------------------------------------------------
    all_holdings: list[dict[str, Any]] = []

    for result in recovered:
        all_holdings.append(
            {
                "excelRow": result["excelRow"],
                "prudentialUrl": result["prudentialUrl"],
                "pruAccessName": result["pruAccessName"],
                "holdings": result["holdings"],
                "verificationPass": result["verificationPass"],
                "parser": result["parser"],
            }
        )

    write_json(
        RECOVERY_ALL_HOLDINGS,
        all_holdings,
    )

    # ------------------------------------------------------------------------
    # Final summary.
    # ------------------------------------------------------------------------
    recovery_status = (
        "success"
        if not failed
        else "partial"
    )

    summary = {
        "status": recovery_status,
        "generatedAtUtc": utc_now(),

        "frozenBaseline": {
            "path": str(BASELINE_SUMMARY),
            "status": baseline.get("status"),
            "successfulFunds": baseline.get("successfulFunds"),
            "failedFunds": baseline.get("failedFunds"),
            "failedFundsDetailCount": len(failed_details),
        },

        "failedFundsFromBaseline": len(failed_details),

        "matchedRecoveryTargets": len(
            recovery_targets
        ),

        "successfulFunds": len(recovered),
        "failedFunds": len(failed),

        "recovered": recovered,
        "failed": failed,

        "rules": {
            "baselineModified": False,
            "pruaccessModified": False,
            "baselineOutputOverwritten": False,
            "officialPrudentialSourceOnly": True,
            "officialFactsheetRequired": True,
            "physicalPdfCoordinatesUsed": True,
            "twoColumnAware": True,
            "multiLineNamesSupported": True,
            "fixedIncomeLastPercentageHandling": True,
            "publishedPercentageRequired": True,
            "nameMustBePresentInOfficialPdf": True,
            "percentageMustBePresentInOfficialPdf": True,
            "namePercentageProximityVerification": True,
            "syntheticData": False,
            "estimatedData": False,
            "fabricatedData": False,
            "interpolatedData": False,
            "forcedTenHoldings": False,
            "maximumHoldings": MAX_TOP_HOLDINGS,
        },
    }

    write_json(
        RECOVERY_SUMMARY,
        summary,
    )

    # ------------------------------------------------------------------------
    # Console summary.
    # ------------------------------------------------------------------------
    print()
    print("=" * 72)
    print("RECOVERY COMPLETE")
    print("=" * 72)

    print()
    print("Status:")
    print(recovery_status)

    print()
    print("Baseline failed funds:")
    print(len(failed_details))

    print()
    print("Matched recovery targets:")
    print(len(recovery_targets))

    print()
    print("Recovery successful:")
    print(len(recovered))

    print()
    print("Recovery failed:")
    print(len(failed))

    print()
    print("Recovered funds:")

    if recovered:
        for result in recovered:
            print(
                f"  - Row {result['excelRow']}: "
                f"{result['pruAccessName']} "
                f"({result['topHoldingsCount']} holdings)"
            )
    else:
        print("  None")

    print()
    print("Still failed:")

    if failed:
        for result in failed:
            print(
                f"  - Row {result['excelRow']}: "
                f"{result['error']}"
            )
    else:
        print("  None")

    print()
    print("Recovery output:")
    print(f"  {RECOVERY_ROOT}")

    print()
    print("Frozen baseline remains untouched.")

    # A partial recovery is a valid test result, not a script crash.
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print()
        print("Recovery interrupted by user.")
        raise SystemExit(130)
    except Exception as exc:
        print()
        print("=" * 72)
        print("FATAL RECOVERY ERROR")
        print("=" * 72)
        print(str(exc))
        print()
        raise SystemExit(1)
