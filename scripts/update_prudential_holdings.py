#!/usr/bin/env python3

"""
VGrat FMS - Prudential Failed-Fund Holdings Recovery TEST

PURPOSE
=======

This script is a SEPARATE recovery/test script for the frozen
test_prudential_holdings.py baseline.

It does NOT modify:

    scripts/test_prudential_holdings.py
    scripts/test_pruaccess.py

It does NOT overwrite:

    output_holdings/

It writes recovery/test results only to:

    output_holdings_recovery/

WORKFLOW
========

GitHub Actions
      |
      v
Frozen baseline:
    test_prudential_holdings.py
      |
      v
output_holdings/run_summary.json
      |
      v
THIS SCRIPT
      |
      v
Read run_summary["failed"]
      |
      v
Read matching URLs from Funds Links.xlsm
      |
      v
Open official Prudential fund page
      |
      v
Locate official factsheet
      |
      v
Download official Prudential PDF
      |
      v
Extract physical PDF coordinates
      |
      v
Locate visual "Top 10 Holdings" section
      |
      v
Reconstruct visual holding rows
      |
      v
Handle detached percentages / multiline names
      |
      v
Validate every recovered holding against the PDF
      |
      v
Re-download official PDF
      |
      v
Run final verification
      |
      v
Save verified recovery result

IMPORTANT
=========

The frozen baseline's run_summary.json uses:

    "failedFunds": 9

as a NUMERIC COUNT.

The actual failed fund records are stored in:

    "failed": [
        {...},
        {...}
    ]

This recovery script therefore reads:

    baseline["failed"]

and validates it against:

    baseline["failedFunds"]

No hardcoded failed-row list is used.

HARD RULES
==========

- Only failed funds from the frozen baseline are processed.
- No successful baseline fund is reinterpreted.
- Excel Column A remains the master fund universe.
- No hardcoded 67-fund limit.
- Only official Prudential Singapore URLs are accepted.
- Only official Prudential Singapore factsheets are accepted.
- No third-party holdings source.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No forced 10 holdings.
- Fewer than 10 published holdings are allowed.
- Published order is preserved where recoverable.
- Duplicate holding names are allowed.
- Duplicate holding percentages are allowed.
- Fixed-income internal percentages are not treated as portfolio weights
  when a later percentage exists on the same visual holding row.
- Detached portfolio percentages may be paired with the visually preceding
  holding name when physical PDF coordinates establish that relationship.
- Multi-line holding names are supported.
- Two-column PDF layouts are separated spatially.
- Every accepted holding must have a published percentage.
- Every accepted holding name and percentage must be verifiable against
  the official PDF text.
- The official PDF is downloaded again for final verification.
- Failed recovery remains failed.
- No blank holdings are inserted.
- No synthetic data is generated.
- No dashboard data.json is created.
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

BASELINE_OUTPUT_DIR = Path(
    "output_holdings"
)

BASELINE_RUN_SUMMARY_FILE = (
    BASELINE_OUTPUT_DIR / "run_summary.json"
)

RECOVERY_OUTPUT_DIR = Path(
    "output_holdings_recovery"
)

RECOVERY_FUNDS_DIR = (
    RECOVERY_OUTPUT_DIR / "funds"
)

RECOVERY_RUN_SUMMARY_FILE = (
    RECOVERY_OUTPUT_DIR / "run_summary.json"
)

RECOVERED_HOLDINGS_FILE = (
    RECOVERY_OUTPUT_DIR / "recovered_holdings.json"
)


# =============================================================================
# BROWSER CONFIGURATION
# =============================================================================

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3

RETRY_DELAY_SECONDS = 3.0


# =============================================================================
# HOLDINGS CONFIGURATION
# =============================================================================

MAX_HOLDINGS = 10

MIN_WEIGHT = 0.000001

MAX_WEIGHT = 100.0


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

def clean_text(value) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = text.replace(
        "\u200b",
        "",
    )

    text = text.replace(
        "\u200c",
        "",
    )

    text = text.replace(
        "\u200d",
        "",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(value) -> str:
    text = clean_text(value)

    text = text.casefold()

    text = text.replace(
        "–",
        "-",
    )

    text = text.replace(
        "—",
        "-",
    )

    text = text.replace(
        "−",
        "-",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


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


# =============================================================================
# PERCENTAGE HELPERS
# =============================================================================

PERCENTAGE_PATTERN = re.compile(
    r"""
    (?<![\d.])
    (?P<number>
        (?:\d+(?:\.\d+)?)
        |
        (?:\.\d+)
    )
    \s*
    %
    """,
    re.VERBOSE,
)


def find_percentages(
    text: str,
) -> list[dict]:

    results = []

    for match in PERCENTAGE_PATTERN.finditer(
        text
    ):

        number_text = match.group(
            "number"
        )

        try:
            number = float(
                number_text
            )
        except ValueError:
            continue

        if number < MIN_WEIGHT:
            continue

        if number > MAX_WEIGHT:
            continue

        results.append(
            {
                "value": number,
                "text": clean_text(
                    match.group(0)
                ),
                "start": match.start(),
                "end": match.end(),
            }
        )

    return results


def find_last_percentage(
    text: str,
) -> dict | None:

    percentages = find_percentages(
        text
    )

    if not percentages:
        return None

    return percentages[-1]


def remove_percentage(
    text: str,
    percentage: dict,
) -> str:

    start = percentage["start"]
    end = percentage["end"]

    result = (
        text[:start]
        + " "
        + text[end:]
    )

    return clean_text(
        result
    )


# =============================================================================
# BASELINE HANDOFF
# =============================================================================

def load_baseline_summary() -> dict:

    print()
    print(
        "=" * 60
    )
    print(
        "READING FROZEN BASELINE"
    )
    print(
        "=" * 60
    )

    if not BASELINE_RUN_SUMMARY_FILE.exists():

        raise FileNotFoundError(
            "Frozen baseline run_summary.json was not found: "
            f"{BASELINE_RUN_SUMMARY_FILE}"
        )

    try:

        data = json.loads(
            BASELINE_RUN_SUMMARY_FILE.read_text(
                encoding="utf-8"
            )
        )

    except Exception as error:

        raise RuntimeError(
            "Could not read frozen baseline run_summary.json: "
            f"{error}"
        ) from error

    if not isinstance(
        data,
        dict,
    ):

        raise RuntimeError(
            "Frozen baseline run_summary.json is not a JSON object."
        )

    failed_count = data.get(
        "failedFunds"
    )

    failed_records = data.get(
        "failed"
    )

    print(
        f"Baseline status: {data.get('status')}"
    )

    print(
        f"Baseline successful funds: "
        f"{data.get('successfulFunds')}"
    )

    print(
        f"Baseline failedFunds count: "
        f"{failed_count}"
    )

    if failed_records is None:

        raise RuntimeError(
            "Frozen baseline does not contain the required "
            "'failed' list."
        )

    if not isinstance(
        failed_records,
        list,
    ):

        raise RuntimeError(
            "Frozen baseline 'failed' field is not a list."
        )

    if not isinstance(
        failed_count,
        int,
    ):

        raise RuntimeError(
            "Frozen baseline 'failedFunds' field is not a numeric count."
        )

    if failed_count != len(
        failed_records
    ):

        raise RuntimeError(
            "Frozen baseline integrity error: "
            f"failedFunds={failed_count} but "
            f"len(failed)={len(failed_records)}."
        )

    if failed_count == 0:

        print()
        print(
            "Baseline reports zero failed funds."
        )

        return data

    print()
    print(
        "Actual failed fund records:"
    )

    for item in failed_records:

        print(
            f" - Row {item.get('excelRow')}: "
            f"{item.get('error', 'Unknown error')}"
        )

    return data


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> dict[int, dict]:

    print()
    print(
        "=" * 60
    )
    print(
        "READING EXCEL MASTER FUND UNIVERSE"
    )
    print(
        "=" * 60
    )

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

    workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated URLs found in Excel Column A."
        )

    print(
        f"Excel fund universe: {len(funds)}"
    )

    return funds


# =============================================================================
# MATCH BASELINE FAILURES TO EXCEL
# =============================================================================

def resolve_failed_funds(
    baseline: dict,
    excel_funds: dict[int, dict],
) -> list[dict]:

    failed_records = baseline.get(
        "failed",
        []
    )

    resolved = []

    for failure in failed_records:

        if not isinstance(
            failure,
            dict,
        ):

            raise RuntimeError(
                "A baseline failed record is not an object."
            )

        excel_row = failure.get(
            "excelRow"
        )

        if not isinstance(
            excel_row,
            int,
        ):

            raise RuntimeError(
                "Baseline failed record has no valid excelRow: "
                f"{failure}"
            )

        excel_fund = excel_funds.get(
            excel_row
        )

        if excel_fund is None:

            raise RuntimeError(
                "Baseline failed fund row "
                f"{excel_row} does not exist in Funds Links.xlsm."
            )

        merged = dict(
            excel_fund
        )

        merged["baselineFailure"] = failure

        resolved.append(
            merged
        )

    return resolved


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

        if (
            "fund factsheet"
            in text_lower
        ):
            score += 300

        elif (
            "factsheet"
            in text_lower
        ):
            score += 200

        if "factsheet" in href_lower:
            score += 150

        if href_lower.endswith(
            ".pdf"
        ):
            score += 50

        if score <= 0:
            continue

        candidates.append(
            {
                "score": score,
                "url": absolute_url,
                "text": anchor_text,
            }
        )

    if not candidates:

        raise RuntimeError(
            "Could not locate an official Prudential factsheet link."
        )

    candidates.sort(
        key=lambda item: (
            -item["score"],
            item["url"],
        )
    )

    selected = candidates[0]

    return ensure_prudential_url(
        selected["url"]
    )


# =============================================================================
# DOWNLOAD OFFICIAL PDF
# =============================================================================

def download_pdf(
    page,
    pdf_url: str,
) -> bytes:

    pdf_url = ensure_prudential_url(
        pdf_url
    )

    last_error = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):

        try:

            response = page.request.get(
                pdf_url,
                timeout=FACTSHEET_TIMEOUT_MS,
            )

            if not response.ok:

                raise RuntimeError(
                    "Official factsheet HTTP request failed: "
                    f"{response.status}"
                )

            content_type = (
                response.headers.get(
                    "content-type",
                    ""
                )
                .lower()
            )

            body = response.body()

            if not body:

                raise RuntimeError(
                    "Official factsheet response was empty."
                )

            # Prudential may occasionally omit application/pdf from the
            # content type. Therefore validate the actual PDF signature too.
            if not (
                body.startswith(
                    b"%PDF"
                )
                or "pdf" in content_type
            ):

                raise RuntimeError(
                    "Official factsheet response does not appear "
                    "to be a PDF."
                )

            return body

        except Exception as error:

            last_error = error

            if attempt < RETRY_COUNT:
                time.sleep(
                    RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        "Could not download official factsheet after "
        f"{RETRY_COUNT} attempts: {last_error}"
    )


# =============================================================================
# FULL PDF TEXT
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

    if page_count <= 0:

        raise RuntimeError(
            "Factsheet PDF contains no pages."
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
                "Could not extract text from official PDF "
                f"page {page_number}: {error}"
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
            "Official factsheet contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


# =============================================================================
# SPATIAL PDF EXTRACTION
# =============================================================================

def extract_positioned_lines(
    pdf_bytes: bytes,
) -> list[dict]:

    reader = PdfReader(
        BytesIO(
            pdf_bytes
        )
    )

    output = []

    for page_number, pdf_page in enumerate(
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

            if text is None:
                return

            raw = str(
                text
            )

            if not raw.strip():
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

            pieces = raw.splitlines()

            if not pieces:
                pieces = [raw]

            for piece_index, piece in enumerate(
                pieces
            ):

                piece = clean_text(
                    piece
                )

                if not piece:
                    continue

                adjusted_y = y

                if piece_index:

                    try:

                        adjusted_y = (
                            y
                            - (
                                float(
                                    font_size or 8
                                )
                                * piece_index
                                * 1.15
                            )
                        )

                    except Exception:

                        adjusted_y = y

                fragments.append(
                    {
                        "x0": x,
                        "x1": x,
                        "y": adjusted_y,
                        "text": piece,
                    }
                )

        try:

            pdf_page.extract_text(
                visitor_text=visitor_text
            )

        except Exception:
            continue

        # Group fragments by visual baseline.
        groups = []

        for fragment in sorted(
            fragments,
            key=lambda item: (
                -item["y"],
                item["x0"],
            ),
        ):

            target = None

            for group in groups:

                if (
                    abs(
                        group["y"]
                        - fragment["y"]
                    )
                    <= 3.0
                ):

                    target = group
                    break

            if target is None:

                target = {
                    "page": page_number,
                    "y": fragment["y"],
                    "fragments": [],
                }

                groups.append(
                    target
                )

            target["fragments"].append(
                fragment
            )

            target["y"] = (
                sum(
                    item["y"]
                    for item
                    in target["fragments"]
                )
                /
                len(
                    target["fragments"]
                )
            )

        for group in groups:

            ordered = sorted(
                group["fragments"],
                key=lambda item: item["x0"],
            )

            line_text = clean_text(
                " ".join(
                    item["text"]
                    for item
                    in ordered
                )
            )

            if not line_text:
                continue

            output.append(
                {
                    "page": page_number,
                    "y": group["y"],
                    "x0": min(
                        item["x0"]
                        for item
                        in ordered
                    ),
                    "x1": max(
                        item["x1"]
                        for item
                        in ordered
                    ),
                    "text": line_text,
                    "fragments": ordered,
                }
            )

    return output


# =============================================================================
# SPATIAL HEADINGS
# =============================================================================

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


def is_holdings_end_marker(
    text: str,
) -> bool:

    normalized = normalize_text(
        text
    )

    markers = (
        "dividend history",
        "distribution history",
        "performance history",
        "past performance",
        "calendar year performance",
        "fund performance",
        "important information",
        "investment objective",
        "investment objective:",
        "risk factors",
        "disclaimer",
        "source:",
    )

    return any(
        normalized.startswith(
            marker
        )
        for marker
        in markers
    )


# =============================================================================
# COLUMN DETECTION
# =============================================================================

def find_right_column_boundary(
    page_lines: list[dict],
    heading: dict,
) -> float:

    heading_x = heading["x0"]
    heading_y = heading["y"]

    candidates = []

    for line in page_lines:

        if abs(
            line["y"]
            - heading_y
        ) > 55:
            continue

        for fragment in line.get(
            "fragments",
            []
        ):

            if fragment["x0"] <= (
                heading_x + 25
            ):
                continue

            normalized = normalize_text(
                fragment["text"]
            )

            if any(
                marker in normalized
                for marker
                in (
                    "dividend history",
                    "distribution history",
                    "performance history",
                    "calendar year",
                    "date",
                    "frequency",
                )
            ):

                candidates.append(
                    fragment["x0"]
                )

    if candidates:

        return max(
            heading_x + 150,
            min(candidates) - 8,
        )

    # Find a large x-coordinate gap immediately below the heading.
    xs = []

    for line in page_lines:

        if not (
            heading_y - 80
            <= line["y"]
            <= heading_y - 4
        ):
            continue

        for fragment in line.get(
            "fragments",
            []
        ):

            if fragment["x0"] > (
                heading_x + 25
            ):

                xs.append(
                    fragment["x0"]
                )

    if xs:

        candidate = min(
            xs
        )

        if candidate > (
            heading_x + 180
        ):

            return candidate - 8

    return max(
        heading_x + 250,
        430.0,
    )


# =============================================================================
# BUILD VISUAL TOP HOLDINGS SECTION
# =============================================================================

def build_visual_section(
    page_lines: list[dict],
    heading: dict,
) -> list[dict]:

    left_edge = max(
        0.0,
        heading["x0"] - 20.0,
    )

    right_edge = find_right_column_boundary(
        page_lines,
        heading,
    )

    selected = []

    # PDF y coordinates normally decrease as we move down the page.
    for line in page_lines:

        if line["y"] >= (
            heading["y"] - 2.0
        ):
            continue

        if line["y"] < (
            heading["y"] - 560.0
        ):
            continue

        fragments = []

        for fragment in line.get(
            "fragments",
            []
        ):

            x0 = float(
                fragment.get(
                    "x0",
                    0.0
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

        text = clean_text(
            " ".join(
                fragment["text"]
                for fragment
                in fragments
            )
        )

        if not text:
            continue

        if is_top_holdings_heading(
            text
        ):
            continue

        if is_holdings_end_marker(
            text
        ):
            break

        selected.append(
            {
                "page": line["page"],
                "y": line["y"],
                "x0": min(
                    item["x0"]
                    for item
                    in fragments
                ),
                "x1": max(
                    item["x0"]
                    for item
                    in fragments
                ),
                "text": text,
                "fragments": fragments,
            }
        )

    selected.sort(
        key=lambda item: (
            -item["y"],
            item["x0"],
        )
    )

    return selected


# =============================================================================
# RANK DETECTION
# =============================================================================

RANK_PATTERN = re.compile(
    r"""
    ^\s*
    (?P<rank>
        [1-9]
        |
        10
    )
    \s*
    [.)\-:]\s*
    """,
    re.VERBOSE,
)


def extract_rank(
    text: str,
) -> tuple[int | None, str]:

    match = RANK_PATTERN.match(
        text
    )

    if not match:

        return (
            None,
            clean_text(text),
        )

    rank = int(
        match.group(
            "rank"
        )
    )

    remainder = clean_text(
        text[
            match.end():
        ]
    )

    return (
        rank,
        remainder,
    )


# =============================================================================
# NAME CLEANING
# =============================================================================

def clean_holding_name(
    text: str,
) -> str:

    text = clean_text(
        text
    )

    # Remove leading bullets.
    text = re.sub(
        r"^[•·▪◦]+\s*",
        "",
        text,
    )

    # Remove table-column labels that can appear in PDF extraction.
    text = re.sub(
        r"^(?:holding|holdings|security)\s*[:\-]?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    return clean_text(
        text
    )


def looks_like_non_holding_text(
    text: str,
) -> bool:

    normalized = normalize_text(
        text
    )

    if not normalized:
        return True

    if normalized in {
        "holding",
        "holdings",
        "security",
        "weight",
        "weights",
        "name",
        "date",
        "frequency",
        "source",
        "source:",
        "total",
        "total:",
    }:
        return True

    if normalized.startswith(
        "source:"
    ):
        return True

    return False


# =============================================================================
# VISUAL HOLDINGS PARSER
# =============================================================================

def parse_visual_holdings(
    lines: list[dict],
) -> tuple[list[dict], str]:

    holdings = []

    pending_name_lines = []

    pending_rank = None

    last_rank_seen = 0

    diagnostics = []

    def add_pending_text(
        value: str,
    ):

        value = clean_holding_name(
            value
        )

        if not value:
            return

        if looks_like_non_holding_text(
            value
        ):
            return

        pending_name_lines.append(
            value
        )

    def make_holding(
        rank: int | None,
        name: str,
        percentage: dict,
    ):

        nonlocal last_rank_seen

        name = clean_holding_name(
            name
        )

        if not name:
            raise RuntimeError(
                "A published holding percentage was found "
                "but no holding name could be recovered "
                "from the visual PDF layout."
            )

        if rank is None:

            rank = (
                last_rank_seen + 1
            )

        if rank < 1 or rank > MAX_HOLDINGS:

            raise RuntimeError(
                "Recovered holding rank is outside the "
                f"1-{MAX_HOLDINGS} range: {rank}"
            )

        if holdings and rank <= last_rank_seen:

            # Some PDFs omit ranks on continuation lines, but an actual
            # repeated/reversed rank is not accepted.
            if rank != last_rank_seen:

                raise RuntimeError(
                    "Recovered holding rank moved backwards "
                    f"from {last_rank_seen} to {rank}."
                )

        holdings.append(
            {
                "rank": rank,
                "name": name,
                "weight": percentage["value"],
                "weightText": percentage["text"],
            }
        )

        last_rank_seen = max(
            last_rank_seen,
            rank,
        )

    for line in lines:

        text = clean_text(
            line["text"]
        )

        if not text:
            continue

        rank, remainder = extract_rank(
            text
        )

        percentages = find_percentages(
            remainder
        )

        if rank is not None:

            # A new rank establishes a new holding. If the previous pending
            # holding has a detached percentage later, it should already have
            # been completed before this rank. Therefore do not silently drop
            # it.
            if (
                pending_name_lines
                and rank > last_rank_seen + 1
            ):

                diagnostics.append(
                    "Non-contiguous visual rank encountered: "
                    f"{rank}"
                )

            if (
                pending_name_lines
                and rank != last_rank_seen
                and holdings
            ):

                # Flush only if this new ranked row itself has a weight.
                # Otherwise preserve the pending text so a detached weight
                # can still complete it.
                if percentages:

                    pending_name_lines = []

            pending_rank = rank

            if percentages:

                # LAST percentage on the visual row is the portfolio weight.
                last_percentage = percentages[-1]

                name_part = remove_percentage(
                    remainder,
                    last_percentage,
                )

                # Remove any earlier percentages only if they are separate
                # table values. For fixed-income holdings, earlier percentages
                # belong to the security name and MUST remain.
                name_part = clean_text(
                    name_part
                )

                if name_part:

                    add_pending_text(
                        name_part
                    )

                name = clean_text(
                    " ".join(
                        pending_name_lines
                    )
                )

                make_holding(
                    pending_rank,
                    name,
                    last_percentage,
                )

                pending_name_lines = []
                pending_rank = None

                if len(
                    holdings
                ) >= MAX_HOLDINGS:

                    break

            else:

                # Rank with no percentage: the name continues across the next
                # visual line(s).
                if remainder:

                    add_pending_text(
                        remainder
                    )

            continue

        # ---------------------------------------------------------------------
        # UNRANKED VISUAL LINE
        # ---------------------------------------------------------------------

        if percentages:

            last_percentage = percentages[-1]

            name_part = remove_percentage(
                text,
                last_percentage,
            )

            name_part = clean_text(
                name_part
            )

            if name_part:

                add_pending_text(
                    name_part
                )

            if pending_name_lines:

                name = clean_text(
                    " ".join(
                        pending_name_lines
                    )
                )

                make_holding(
                    pending_rank,
                    name,
                    last_percentage,
                )

                pending_name_lines = []
                pending_rank = None

                if len(
                    holdings
                ) >= MAX_HOLDINGS:

                    break

            continue

        # No percentage. Treat as continuation of the current visual holding.
        if text:

            add_pending_text(
                text
            )

    # If text remains after the final row, it means the PDF published a
    # holding name without a published percentage. This must not be accepted.
    if pending_name_lines:

        diagnostics.append(
            "Unresolved holding-name text remained after "
            "the final visual row."
        )

    if not holdings:

        detail = (
            "; ".join(
                diagnostics
            )
            if diagnostics
            else "No visual holdings were recovered."
        )

        raise RuntimeError(
            "Spatial holdings parser failed: "
            f"{detail}"
        )

    # Require sequential published order.
    expected = 1

    for item in holdings:

        if item["rank"] != expected:

            raise RuntimeError(
                "Recovered visual holding ranks are not in published order: "
                f"expected {expected}, found {item['rank']}."
            )

        expected += 1

    return (
        holdings,
        "\n".join(
            diagnostics
        ),
    )


# =============================================================================
# ALTERNATIVE VISUAL PAIRING
# =============================================================================

def parse_visual_holdings_with_detached_weight_pairing(
    lines: list[dict],
) -> list[dict]:

    """
    More permissive spatial parser for PDFs where the portfolio weight is
    printed on a physically separate line.

    Example:

        1. Microsoft Corporation
           United States
           8.9%

        2. Alphabet Inc.
           Class A
           5.4%

    The parser uses the visual sequence and does not use external information.
    """

    rows = []

    current = None

    for line in lines:

        text = clean_text(
            line["text"]
        )

        if not text:
            continue

        rank, remainder = extract_rank(
            text
        )

        if rank is not None:

            if current is not None:

                rows.append(
                    current
                )

            current = {
                "rank": rank,
                "name_lines": [],
                "weight": None,
                "source_lines": [],
            }

            current["source_lines"].append(
                text
            )

            percentages = find_percentages(
                remainder
            )

            if percentages:

                last_percentage = percentages[-1]

                name_part = remove_percentage(
                    remainder,
                    last_percentage,
                )

                if clean_text(
                    name_part
                ):

                    current["name_lines"].append(
                        clean_text(
                            name_part
                        )
                    )

                current["weight"] = (
                    last_percentage
                )

            else:

                if remainder:

                    current["name_lines"].append(
                        clean_text(
                            remainder
                        )
                    )

            continue

        if current is None:
            continue

        current["source_lines"].append(
            text
        )

        percentages = find_percentages(
            text
        )

        if percentages:

            last_percentage = percentages[-1]

            # If the line contains text as well as a percentage, retain the
            # text because it can be a continuation of a fixed-income name.
            name_part = remove_percentage(
                text,
                last_percentage,
            )

            if clean_text(
                name_part
            ):

                current["name_lines"].append(
                    clean_text(
                        name_part
                    )
                )

            current["weight"] = (
                last_percentage
            )

        else:

            if (
                current["weight"]
                is None
            ):

                if not looks_like_non_holding_text(
                    text
                ):

                    current["name_lines"].append(
                        text
                    )

    if current is not None:

        rows.append(
            current
        )

    holdings = []

    for row in rows:

        if row["weight"] is None:
            continue

        name = clean_text(
            " ".join(
                row["name_lines"]
            )
        )

        if not name:
            continue

        holdings.append(
            {
                "rank": row["rank"],
                "name": clean_holding_name(
                    name
                ),
                "weight": row["weight"]["value"],
                "weightText": row["weight"]["text"],
            }
        )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    if not holdings:

        raise RuntimeError(
            "Detached-weight spatial parser recovered "
            "no complete holding rows."
        )

    expected = 1

    for item in holdings:

        if item["rank"] != expected:

            raise RuntimeError(
                "Detached-weight parser recovered non-sequential ranks: "
                f"expected {expected}, found {item['rank']}."
            )

        expected += 1

    return holdings


# =============================================================================
# HOLDINGS HEADING SEARCH
# =============================================================================

def recover_from_spatial_pdf(
    pdf_bytes: bytes,
) -> tuple[list[dict], str, dict]:

    positioned = extract_positioned_lines(
        pdf_bytes
    )

    if not positioned:

        raise RuntimeError(
            "No positioned PDF text was available."
        )

    candidates = []

    page_numbers = sorted(
        {
            line["page"]
            for line
            in positioned
        }
    )

    for page_number in page_numbers:

        page_lines = [
            line
            for line
            in positioned
            if line["page"] == page_number
        ]

        headings = [
            line
            for line
            in page_lines
            if is_top_holdings_heading(
                line["text"]
            )
        ]

        for heading in headings:

            visual_lines = build_visual_section(
                page_lines,
                heading,
            )

            if not visual_lines:
                continue

            # First parser: strict visual grouping.
            try:

                holdings, diagnostics = parse_visual_holdings(
                    visual_lines
                )

                candidates.append(
                    {
                        "page": page_number,
                        "heading": heading["text"],
                        "holdings": holdings,
                        "diagnostics": diagnostics,
                        "parser": "spatial-strict",
                        "visualLines": visual_lines,
                    }
                )

            except Exception as strict_error:

                # Second parser: explicitly designed for detached weights.
                try:

                    holdings = parse_visual_holdings_with_detached_weight_pairing(
                        visual_lines
                    )

                    candidates.append(
                        {
                            "page": page_number,
                            "heading": heading["text"],
                            "holdings": holdings,
                            "diagnostics": (
                                "Strict spatial parser failed: "
                                f"{clean_text(str(strict_error))}"
                            ),
                            "parser": "spatial-detached-weight",
                            "visualLines": visual_lines,
                        }
                    )

                except Exception:
                    continue

    if not candidates:

        raise RuntimeError(
            "No usable official Top 10 Holdings visual section "
            "could be recovered from the PDF."
        )

    # Prefer the candidate with the greatest number of published rows.
    # Ties are resolved by earlier page occurrence.
    candidates.sort(
        key=lambda candidate: (
            -len(
                candidate["holdings"]
            ),
            candidate["page"],
        )
    )

    selected = candidates[0]

    return (
        selected["holdings"],
        selected["heading"],
        selected,
    )


# =============================================================================
# PDF VERIFICATION
# =============================================================================

def percentage_text_variants(
    weight: float,
    weight_text: str,
) -> list[str]:

    variants = []

    if weight_text:
        variants.append(
            normalize_text(
                weight_text
            )
        )

    # Use only textual variants actually representing the same published
    # numeric percentage. This is verification normalization, not fabrication.
    variants.append(
        normalize_text(
            f"{weight:g}%"
        )
    )

    variants.append(
        normalize_text(
            f"{weight:.1f}%"
        )
    )

    variants.append(
        normalize_text(
            f"{weight:.2f}%"
        )
    )

    return list(
        dict.fromkeys(
            variants
        )
    )


def verify_holding_against_pdf(
    holding: dict,
    full_pdf_text: str,
) -> dict:

    normalized_pdf = normalize_text(
        full_pdf_text
    )

    name = normalize_text(
        holding["name"]
    )

    if not name:

        raise RuntimeError(
            "Recovered holding has an empty name."
        )

    # Names may contain line-wrap hyphens. Verification first tries the exact
    # normalized name, then a de-hyphenated representation.
    name_found = (
        name in normalized_pdf
    )

    if not name_found:

        dehyphenated_name = (
            name.replace(
                "- ",
                "",
            )
        )

        dehyphenated_pdf = (
            normalized_pdf.replace(
                "- ",
                "",
            )
        )

        name_found = (
            dehyphenated_name
            in dehyphenated_pdf
        )

    if not name_found:

        raise RuntimeError(
            "Recovered holding name could not be verified "
            "against the official PDF: "
            f"{holding['name']}"
        )

    weight = float(
        holding["weight"]
    )

    variants = percentage_text_variants(
        weight,
        holding.get(
            "weightText",
            ""
        ),
    )

    weight_found = any(
        variant
        in normalized_pdf
        for variant
        in variants
    )

    if not weight_found:

        raise RuntimeError(
            "Recovered holding percentage could not be verified "
            "against the official PDF: "
            f"{holding.get('weightText')}"
        )

    return {
        "nameVerified": True,
        "weightVerified": True,
    }


def verify_all_holdings(
    holdings: list[dict],
    full_pdf_text: str,
) -> list[dict]:

    if not holdings:

        raise RuntimeError(
            "No holdings available for verification."
        )

    verification = []

    for holding in holdings:

        result = verify_holding_against_pdf(
            holding,
            full_pdf_text,
        )

        verification.append(
            {
                "rank": holding["rank"],
                "name": holding["name"],
                "weightText": holding["weightText"],
                **result,
            }
        )

    return verification


# =============================================================================
# SECOND-DOWNLOAD FINAL VERIFICATION
# =============================================================================

def final_redownload_verification(
    page,
    factsheet_url: str,
    holdings: list[dict],
) -> dict:

    second_pdf = download_pdf(
        page,
        factsheet_url,
    )

    second_text, second_page_count = extract_pdf_text(
        second_pdf
    )

    verification = verify_all_holdings(
        holdings,
        second_text,
    )

    return {
        "pdfBytes": second_pdf,
        "pdfText": second_text,
        "pageCount": second_page_count,
        "verification": verification,
    }


# =============================================================================
# SAVE RECOVERY FUND
# =============================================================================

def save_recovery_fund(
    fund: dict,
    factsheet_url: str,
    pdf_bytes: bytes,
    pdf_text: str,
    heading: str,
    holdings: list[dict],
    verification: list[dict],
    parser_name: str,
    source_metadata: dict,
) -> Path:

    row = fund["excelRow"]

    fund_name = (
        source_metadata.get(
            "fundName"
        )
        or fund.get(
            "pruAccessName"
        )
        or f"row_{row}"
    )

    directory = (
        RECOVERY_FUNDS_DIR
        /
        f"{row}_{safe_filename(fund_name)}"
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

    top_section_text = "\n".join(
        [
            f"{item['rank']}. "
            f"{item['name']} "
            f"- {item['weightText']}"
            for item
            in holdings
        ]
    )

    (directory / "top_holdings_visual_section.txt").write_text(
        top_section_text,
        encoding="utf-8",
    )

    save_json(
        directory / "top_holdings.json",
        {
            "source": "Prudential Singapore",
            "parser": parser_name,
            "heading": heading,
            "count": len(
                holdings
            ),
            "holdings": holdings,
            "verification": verification,
        },
    )

    save_json(
        directory / "metadata.json",
        {
            "excelRow": row,
            "prudentialUrl": fund["prudentialUrl"],
            "pruAccessName": fund.get(
                "pruAccessName"
            ),
            "factsheetUrl": factsheet_url,
            "fundName": source_metadata.get(
                "fundName"
            ),
            "factsheetDocumentDate": source_metadata.get(
                "factsheetDocumentDate"
            ),
            "factsheetDataAsAt": source_metadata.get(
                "factsheetDataAsAt"
            ),
            "parser": parser_name,
            "topHoldingsCount": len(
                holdings
            ),
            "finalVerificationPassed": True,
        },
    )

    return directory


# =============================================================================
# FUND PAGE METADATA
# =============================================================================

def extract_page_metadata(
    page,
) -> dict:

    result = {
        "fundName": None,
        "factsheetDocumentDate": None,
        "factsheetDataAsAt": None,
    }

    try:

        body_text = clean_text(
            page.locator(
                "body"
            ).inner_text()
        )

    except Exception:

        return result

    # Best-effort metadata only. It never determines holding values.
    lines = [
        clean_text(
            line
        )
        for line
        in body_text.splitlines()
        if clean_text(line)
    ]

    for line in lines[:80]:

        normalized = normalize_text(
            line
        )

        if (
            "fund name"
            in normalized
            and ":"
            in line
        ):

            result["fundName"] = clean_text(
                line.split(
                    ":",
                    1
                )[1]
            )

            break

    return result


# =============================================================================
# PROCESS ONE FUND
# =============================================================================

def process_failed_fund(
    page,
    fund: dict,
) -> dict:

    row = fund["excelRow"]

    prudential_url = ensure_prudential_url(
        fund["prudentialUrl"]
    )

    print()
    print(
        "-" * 70
    )

    print(
        f"RECOVERING EXCEL ROW {row}"
    )

    print(
        f"Prudential URL: {prudential_url}"
    )

    print(
        f"PruAccess reference: "
        f"{fund.get('pruAccessName') or '-'}"
    )

    last_error = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):

        try:

            print(
                f"Recovery attempt {attempt}/{RETRY_COUNT}"
            )

            page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            page.wait_for_timeout(
                POST_PAGE_WAIT_MS
            )

            # Confirm final page remains official Prudential.
            if not is_prudential_url(
                page.url
            ):

                raise RuntimeError(
                    "Prudential page redirected to a non-Prudential host: "
                    f"{page.url}"
                )

            source_metadata = extract_page_metadata(
                page
            )

            factsheet_url = find_factsheet_url(
                page
            )

            print(
                f"Factsheet: {factsheet_url}"
            )

            pdf_bytes = download_pdf(
                page,
                factsheet_url,
            )

            pdf_text, page_count = extract_pdf_text(
                pdf_bytes
            )

            holdings, heading, candidate = recover_from_spatial_pdf(
                pdf_bytes
            )

            print(
                f"Spatial parser: "
                f"{candidate['parser']}"
            )

            print(
                f"Top Holdings recovered: "
                f"{len(holdings)}"
            )

            for holding in holdings:

                print(
                    f"  {holding['rank']}. "
                    f"{holding['name']} "
                    f"- {holding['weightText']}"
                )

            first_verification = verify_all_holdings(
                holdings,
                pdf_text,
            )

            print(
                "Initial PDF verification: PASS"
            )

            # -----------------------------------------------------------------
            # FINAL RE-DOWNLOAD
            # -----------------------------------------------------------------

            final_check = final_redownload_verification(
                page,
                factsheet_url,
                holdings,
            )

            print(
                "Final PDF re-download verification: PASS"
            )

            output_directory = save_recovery_fund(
                fund=fund,
                factsheet_url=factsheet_url,
                pdf_bytes=final_check["pdfBytes"],
                pdf_text=final_check["pdfText"],
                heading=heading,
                holdings=holdings,
                verification=final_check["verification"],
                parser_name=candidate["parser"],
                source_metadata=source_metadata,
            )

            return {
                "status": "success",
                "excelRow": row,
                "prudentialUrl": prudential_url,
                "pruAccessName": fund.get(
                    "pruAccessName"
                ),
                "fundName": (
                    source_metadata.get(
                        "fundName"
                    )
                    or fund.get(
                        "pruAccessName"
                    )
                ),
                "factsheetUrl": factsheet_url,
                "topHoldingsCount": len(
                    holdings
                ),
                "parser": candidate["parser"],
                "heading": heading,
                "pageCount": page_count,
                "initialVerificationPassed": True,
                "finalVerificationPassed": True,
                "outputDirectory": str(
                    output_directory
                ),
                "holdings": holdings,
            }

        except (
            PlaywrightTimeoutError,
            Exception,
        ) as error:

            last_error = clean_text(
                str(error)
            )

            print(
                f"Recovery attempt failed: {last_error}"
            )

            if attempt < RETRY_COUNT:

                time.sleep(
                    RETRY_DELAY_SECONDS
                )

    return {
        "status": "failed",
        "excelRow": row,
        "prudentialUrl": prudential_url,
        "pruAccessName": fund.get(
            "pruAccessName"
        ),
        "error": (
            last_error
            or "Unknown recovery error."
        ),
    }


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    print()
    print(
        "=" * 60
    )
    print(
        "VGrat FMS - PRUDENTIAL FAILED-FUND RECOVERY TEST"
    )
    print(
        "=" * 60
    )

    print()
    print(
        "IMPORTANT:"
    )

    print(
        "Frozen baseline will NOT be modified."
    )

    print(
        "PruAccess script will NOT be modified."
    )

    print(
        "Baseline output will NOT be overwritten."
    )

    print(
        "Recovery output:"
    )

    print(
        f"  {RECOVERY_OUTPUT_DIR}"
    )

    print(
        "=" * 60
    )

    started_at = utc_now_iso()

    try:

        baseline = load_baseline_summary()

        excel_funds = read_excel_funds()

        failed_funds = resolve_failed_funds(
            baseline,
            excel_funds,
        )

        expected_failed_count = baseline.get(
            "failedFunds"
        )

        print()
        print(
            "=" * 60
        )
        print(
            "RECOVERY TARGETS"
        )
        print(
            "=" * 60
        )

        print(
            f"Baseline failed-fund count: "
            f"{expected_failed_count}"
        )

        print(
            f"Resolved recovery targets: "
            f"{len(failed_funds)}"
        )

        if len(failed_funds) != expected_failed_count:

            raise RuntimeError(
                "Recovery target count mismatch: "
                f"baseline={expected_failed_count}, "
                f"resolved={len(failed_funds)}."
            )

        # ---------------------------------------------------------------------
        # ZERO FAILURES
        # ---------------------------------------------------------------------

        if not failed_funds:

            completed_at = utc_now_iso()

            summary = {
                "status": "success",
                "startedAtUtc": started_at,
                "completedAtUtc": completed_at,
                "baselineStatus": baseline.get(
                    "status"
                ),
                "failedFundsFromBaseline": 0,
                "matchedRecoveryTargets": 0,
                "successfulFunds": 0,
                "failedFunds": 0,
                "recovered": [],
                "failed": [],
                "rules": {
                    "frozenBaselineUntouched": True,
                    "baselineFailureListReadFromFailedField": True,
                    "dynamicFailedFundSelection": True,
                    "officialPrudentialOnly": True,
                    "officialFactsheetOnly": True,
                    "spatialPdfRecovery": True,
                    "noInferredHoldings": True,
                    "noFabricatedData": True,
                    "noSyntheticData": True,
                    "noForcedTenHoldings": True,
                    "finalRedownloadVerificationRequired": True,
                    "baselineOutputUntouched": True,
                    "pruAccessScriptUntouched": True,
                },
            }

            save_json(
                RECOVERY_RUN_SUMMARY_FILE,
                summary,
            )

            save_json(
                RECOVERED_HOLDINGS_FILE,
                {
                    "status": "success",
                    "recovered": [],
                },
            )

            print()
            print(
                "No failed funds reported by the frozen baseline."
            )

            return 0

        successful = []
        failed = []

        # ---------------------------------------------------------------------
        # PLAYWRIGHT
        # ---------------------------------------------------------------------

        with sync_playwright() as playwright:

            browser = playwright.chromium.launch(
                headless=BROWSER_HEADLESS,
            )

            context = browser.new_context(
                ignore_https_errors=False,
                accept_downloads=True,
            )

            page = context.new_page()

            page.set_default_timeout(
                PAGE_TIMEOUT_MS
            )

            try:

                for index, fund in enumerate(
                    failed_funds,
                    start=1,
                ):

                    print()
                    print(
                        "=" * 70
                    )

                    print(
                        f"RECOVERY TARGET "
                        f"{index}/{len(failed_funds)}"
                    )

                    print(
                        f"Excel row: "
                        f"{fund['excelRow']}"
                    )

                    print(
                        f"Fund: "
                        f"{fund.get('pruAccessName') or '-'}"
                    )

                    print(
                        "=" * 70
                    )

                    result = process_failed_fund(
                        page,
                        fund,
                    )

                    if result.get(
                        "status"
                    ) == "success":

                        successful.append(
                            result
                        )

                        print()
                        print(
                            f"STATUS: RECOVERED "
                            f"(row {fund['excelRow']})"
                        )

                    else:

                        failed.append(
                            result
                        )

                        print()
                        print(
                            f"STATUS: STILL FAILED "
                            f"(row {fund['excelRow']})"
                        )

            finally:

                context.close()
                browser.close()

        # ---------------------------------------------------------------------
        # FINAL SUMMARY
        # ---------------------------------------------------------------------

        completed_at = utc_now_iso()

        summary = {
            "status": (
                "success"
                if not failed
                else "partial"
            ),
            "startedAtUtc": started_at,
            "completedAtUtc": completed_at,
            "baselineStatus": baseline.get(
                "status"
            ),
            "excelFundUniverse": len(
                excel_funds
            ),
            "failedFundsFromBaseline": len(
                failed_funds
            ),
            "matchedRecoveryTargets": len(
                failed_funds
            ),
            "successfulFunds": len(
                successful
            ),
            "failedFunds": len(
                failed
            ),
            "recovered": [
                {
                    "excelRow": item.get(
                        "excelRow"
                    ),
                    "fundName": item.get(
                        "fundName"
                    ),
                    "factsheetUrl": item.get(
                        "factsheetUrl"
                    ),
                    "parser": item.get(
                        "parser"
                    ),
                    "heading": item.get(
                        "heading"
                    ),
                    "topHoldingsCount": item.get(
                        "topHoldingsCount"
                    ),
                    "finalVerificationPassed": item.get(
                        "finalVerificationPassed"
                    ),
                }
                for item
                in successful
            ],
            "failed": failed,
            "rules": {
                "frozenBaselineUntouched": True,
                "baselineFailureListReadFromFailedField": True,
                "failedFundsIsCountOnly": True,
                "dynamicFailedFundSelection": True,
                "excelColumnAControlsUniverse": True,
                "officialPrudentialOnly": True,
                "officialFactsheetOnly": True,
                "spatialPdfRecovery": True,
                "detachedPercentageRecovery": True,
                "multilineHoldingRecovery": True,
                "twoColumnLayoutSeparation": True,
                "fixedIncomeLastPercentageRule": True,
                "publishedOrderRequired": True,
                "noInferredHoldings": True,
                "noFabricatedData": True,
                "noSyntheticData": True,
                "noForcedTenHoldings": True,
                "fewerThanTenAllowed": True,
                "duplicateHoldingNamesAllowed": True,
                "duplicateHoldingPercentagesAllowed": True,
                "initialPdfVerificationRequired": True,
                "finalRedownloadVerificationRequired": True,
                "baselineOutputUntouched": True,
                "pruAccessScriptUntouched": True,
                "dashboardDataUntouched": True,
            },
        }

        save_json(
            RECOVERY_RUN_SUMMARY_FILE,
            summary,
        )

        save_json(
            RECOVERED_HOLDINGS_FILE,
            {
                "status": summary["status"],
                "source": "Prudential Singapore",
                "generatedAtUtc": completed_at,
                "baselineFailedFunds": len(
                    failed_funds
                ),
                "successfulFunds": len(
                    successful
                ),
                "failedFunds": len(
                    failed
                ),
                "funds": successful,
            },
        )

        # ---------------------------------------------------------------------
        # CONSOLE SUMMARY
        # ---------------------------------------------------------------------

        print()
        print(
            "=" * 70
        )

        print(
            "PRUDENTIAL HOLDINGS RECOVERY SUMMARY"
        )

        print(
            "=" * 70
        )

        print(
            f"Baseline failures: "
            f"{len(failed_funds)}"
        )

        print(
            f"Recovery successful: "
            f"{len(successful)}"
        )

        print(
            f"Recovery still failed: "
            f"{len(failed)}"
        )

        print()

        print(
            "Recovered funds:"
        )

        for item in successful:

            print(
                f" - Row {item.get('excelRow')}: "
                f"{item.get('fundName') or '-'} "
                f"({item.get('topHoldingsCount')} holdings)"
            )

        print()

        print(
            "Still failed:"
        )

        for item in failed:

            print(
                f" - Row {item.get('excelRow')}: "
                f"{item.get('error', 'Unknown error')}"
            )

        print()

        print(
            f"Recovery summary: "
            f"{RECOVERY_RUN_SUMMARY_FILE}"
        )

        print(
            f"Recovered holdings: "
            f"{RECOVERED_HOLDINGS_FILE}"
        )

        print(
            "=" * 70
        )

        if failed:

            print()
            print(
                "RECOVERY TEST RESULT: PARTIAL"
            )

            print(
                "One or more baseline failures remain unresolved."
            )

            return 1

        print()
        print(
            "RECOVERY TEST RESULT: SUCCESS"
        )

        print(
            "All baseline failed funds were recovered and "
            "passed final official-PDF verification."
        )

        return 0

    except Exception as error:

        completed_at = utc_now_iso()

        fatal_summary = {
            "status": "fatal_error",
            "startedAtUtc": started_at,
            "completedAtUtc": completed_at,
            "error": clean_text(
                str(error)
            ),
            "rules": {
                "frozenBaselineUntouched": True,
                "baselineOutputUntouched": True,
                "pruAccessScriptUntouched": True,
                "dashboardDataUntouched": True,
            },
        }

        try:

            save_json(
                RECOVERY_RUN_SUMMARY_FILE,
                fatal_summary,
            )

        except Exception:
            pass

        print()
        print(
            "=" * 70
        )

        print(
            "FATAL RECOVERY ERROR:"
        )

        print(
            clean_text(
                str(error)
            )
        )

        print(
            "=" * 70
        )

        return 1


if __name__ == "__main__":
    sys.exit(
        main()
    )
