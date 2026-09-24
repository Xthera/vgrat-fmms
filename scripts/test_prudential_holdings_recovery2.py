#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings RECOVERY 2
ADAPTIVE / FORENSIC EXTRACTOR

RECOVERY SOURCE
===============

Recovery 1:
    output_holdings_recovery/run_summary.json

Recovery 2 ONLY processes funds that failed in Recovery 1.

If Recovery 1's run_summary.json does not contain usable failed-fund
details, Recovery 2 reconstructs the failed-fund list ONLY from:

    output_holdings_recovery/funds/*_failed/failure.json

Recovery 2 NEVER reads:

    output_holdings/run_summary.json

Recovery 2 NEVER scans:

    output_holdings/

The Excel workbook remains authoritative for:
    Column A = Prudential URL
    Column B = PruAccess fund name


PURPOSE
=======

This version is deliberately diagnostic-first.

The previous Recovery 2 implementation stopped whenever it could not find
a literal "Top 10 Holdings" heading. That assumption is not reliable for
all Prudential factsheet PDF layouts.

This version therefore:

1. Loads ONLY Recovery 1 failed funds.
2. Downloads the official Prudential factsheet.
3. Saves the complete PDF and extracted text.
4. Extracts page-by-page text.
5. Extracts positioned PDF text with coordinates.
6. Searches the entire PDF for holdings-related terminology.
7. Searches the entire PDF for ranked candidates 1-10.
8. Searches the entire PDF for published percentages.
9. Searches for conventional holdings sections if present.
10. Builds multiple candidate regions without requiring a heading.
11. Attempts strict text reconstruction.
12. Attempts strict spatial reconstruction.
13. Never invents a holding or percentage.
14. Never pairs arbitrary names and percentages merely because they are
    geographically close.
15. Performs exact re-download verification for any accepted result.
16. Saves forensic diagnostics for every unresolved fund.

IMPORTANT
=========

This script is NOT permitted to conclude that a fund has no holdings merely
because a literal "Top 10 Holdings" heading cannot be found.

"NO_EXPLICIT_SECTION" means only that the conventional heading was not found.

Official PDF evidence is retained for the next extraction stage.


HARD RULES
==========

- Recovery 1 failed funds control Recovery 2 universe.
- No baseline fallback.
- Excel Column A controls the master fund URL.
- No hardcoded master fund count.
- Only official Prudential Singapore pages.
- Only official Prudential Singapore factsheets.
- No third-party holdings sources.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No forced 10 holdings.
- Fewer than 10 published holdings is valid.
- Published holding order is preserved.
- Duplicate holding names are allowed.
- Duplicate holding percentages are allowed.
- Multiline holding names are supported.
- Hyphenated PDF line wraps are reconstructed.
- Fixed-income coupon/rate percentages must not be treated as portfolio
  weights unless the PDF's table structure explicitly identifies the
  percentage as the holding weight.
- No arbitrary coordinate proximity pairing.
- No fuzzy holding matching.
- No synthetic values.
- No estimated values.
- No interpolated values.
- Final verification requires exact rank/name/weight signature equality.
- Recovery 2 writes only to:
      output_holdings_recovery_2/
- Baseline and Recovery 1 outputs are never overwritten.
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
from playwright.sync_api import sync_playwright


# =============================================================================
# CONFIGURATION
# =============================================================================

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

HOLDINGS_KEYWORDS = (
    "holding",
    "holdings",
    "investment",
    "investments",
    "portfolio",
    "security",
    "securities",
    "issuer",
    "issuers",
    "equity",
    "bond",
    "bonds",
    "stock",
    "stocks",
    "top 10",
    "top ten",
)

SECTION_HEADINGS = (
    re.compile(
        r"^top\s+(?:10|ten)\s+holdings\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^top\s+holdings\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^top\s+(?:10|ten)\s+investments\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^largest\s+holdings\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^portfolio\s+holdings\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^major\s+holdings\b",
        re.IGNORECASE,
    ),
)

PERCENTAGE_RE = re.compile(
    r"(?<![\d.])"
    r"(\d+(?:\.\d+)?)"
    r"\s*%"
)

LEADING_RANK_RE = re.compile(
    r"^\s*(\d{1,2})[.)]?\s+(.*)$"
)

RANK_ONLY_RE = re.compile(
    r"^\s*(\d{1,2})[.)]?\s*$"
)

DATE_RE = re.compile(
    r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b"
)


# =============================================================================
# GENERIC HELPERS
# =============================================================================

def utc_now_iso() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


def clean_text(value) -> str:
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def normalize_text(value) -> str:
    text = clean_text(value).lower()

    text = (
        text
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u00a0", " ")
    )

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def save_json(
    path: Path,
    value,
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        json.dumps(
            value,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def safe_filename(
    value: str,
) -> str:

    text = clean_text(value)

    text = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        text,
    )

    text = re.sub(
        r"\s+",
        "_",
        text,
    )

    text = re.sub(
        r"_+",
        "_",
        text,
    )

    return text.strip("._")


def is_prudential_url(
    url: str,
) -> bool:

    try:

        parsed = urlparse(url)

        return (
            parsed.scheme in {"http", "https"}
            and (
                parsed.hostname or ""
            ).lower() in PRUDENTIAL_HOSTS
        )

    except Exception:

        return False


def ensure_prudential_url(
    url: str,
) -> str:

    value = clean_text(url)

    if not value:

        raise RuntimeError(
            "Prudential URL is empty."
        )

    if not is_prudential_url(value):

        raise RuntimeError(
            "URL is not an official Prudential Singapore URL: "
            f"{value}"
        )

    return value


class HoldingsParseFailure(
    RuntimeError
):
    pass


# =============================================================================
# RECOVERY 1 FAILURE LOADING
# =============================================================================

def _normalize_failed_fund_entries(
    candidate,
) -> list[dict]:

    if isinstance(candidate, list):

        entries = candidate

    elif isinstance(candidate, dict):

        entries = []

        for key, value in candidate.items():

            if isinstance(value, dict):

                entry = dict(value)

                if "excelRow" not in entry:

                    for row_key in (
                        "row",
                        "rowNumber",
                        "excel_row",
                        "ExcelRow",
                    ):

                        if row_key in entry:

                            try:

                                entry["excelRow"] = int(
                                    entry[row_key]
                                )

                            except Exception:
                                pass

                            break

                if "excelRow" not in entry:

                    try:

                        entry["excelRow"] = int(key)

                    except Exception:
                        pass

                entries.append(entry)

            elif isinstance(value, str):

                entry = {
                    "error": clean_text(value)
                }

                try:

                    entry["excelRow"] = int(key)

                except Exception:
                    pass

                entries.append(entry)

    else:

        return []

    normalized = []

    for entry in entries:

        if not isinstance(
            entry,
            dict,
        ):

            continue

        item = dict(entry)

        row = None

        for key in (
            "excelRow",
            "row",
            "rowNumber",
            "excel_row",
            "ExcelRow",
        ):

            if key not in item:
                continue

            try:

                row = int(
                    item[key]
                )

            except Exception:

                row = None

            if row is not None:
                break

        if row is None:
            continue

        item["excelRow"] = row

        normalized.append(item)

    normalized.sort(
        key=lambda item: item["excelRow"]
    )

    return normalized


def load_recovery_1_failure_files() -> list[dict]:

    recovery1_funds_dir = (
        Path("output_holdings_recovery")
        / "funds"
    )

    if not recovery1_funds_dir.exists():
        return []

    failure_files = sorted(
        recovery1_funds_dir.glob(
            "*_failed/failure.json"
        ),
        key=lambda path: path.parent.name,
    )

    recovered = []

    for failure_file in failure_files:

        try:

            payload = json.loads(
                failure_file.read_text(
                    encoding="utf-8"
                )
            )

        except Exception as error:

            raise RuntimeError(
                "Could not read Recovery 1 failure file: "
                f"{failure_file}"
            ) from error

        if not isinstance(
            payload,
            dict,
        ):

            continue

        if normalize_text(
            payload.get("status")
        ) != "failed":

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

        if excel_row is None:

            raise RuntimeError(
                "Recovery 1 failure file has no identifiable "
                f"Excel row: {failure_file}"
            )

        try:

            excel_row = int(
                excel_row
            )

        except Exception as error:

            raise RuntimeError(
                "Invalid Excel row in Recovery 1 failure file: "
                f"{failure_file}"
            ) from error

        item = dict(payload)

        item["excelRow"] = excel_row

        item[
            "_reconstructedFromRecovery1Output"
        ] = True

        item[
            "_recovery1FailureFile"
        ] = str(
            failure_file
        )

        recovered.append(item)

    recovered.sort(
        key=lambda item: item["excelRow"]
    )

    return recovered


def load_recovery_1_run_summary() -> dict:

    if not RECOVERY_1_RUN_SUMMARY_FILE.exists():

        raise FileNotFoundError(
            "Recovery 1 run summary not found: "
            f"{RECOVERY_1_RUN_SUMMARY_FILE}"
        )

    try:

        payload = json.loads(
            RECOVERY_1_RUN_SUMMARY_FILE.read_text(
                encoding="utf-8"
            )
        )

    except Exception as error:

        raise RuntimeError(
            "Could not read Recovery 1 run summary: "
            f"{RECOVERY_1_RUN_SUMMARY_FILE}"
        ) from error

    if not isinstance(
        payload,
        dict,
    ):

        raise RuntimeError(
            "Recovery 1 run summary is not a JSON object."
        )

    normalized_failed = (
        _normalize_failed_fund_entries(
            payload.get(
                "failedFundsDetail"
            )
        )
    )

    source = ""

    if normalized_failed:

        source = (
            "recovery1_run_summary:"
            "failedFundsDetail"
        )

    if not normalized_failed:

        alternative_keys = (
            "failedFundsDetails",
            "failedFunds",
            "failedFundsList",
            "failureDetails",
            "failures",
            "failed",
            "failedFundDetails",
        )

        for key in alternative_keys:

            candidate = payload.get(
                key
            )

            candidate_entries = (
                _normalize_failed_fund_entries(
                    candidate
                )
            )

            if candidate_entries:

                normalized_failed = (
                    candidate_entries
                )

                source = (
                    "recovery1_run_summary:"
                    + key
                )

                break

    if not normalized_failed:

        reconstructed = (
            load_recovery_1_failure_files()
        )

        if reconstructed:

            normalized_failed = (
                reconstructed
            )

            source = (
                "recovery1_output_failure_files"
            )

            payload[
                "_recovery2FailureSourceFiles"
            ] = [
                item.get(
                    "_recovery1FailureFile"
                )
                for item in reconstructed
            ]

            print()
            print(
                "WARNING: Recovery 1 run_summary.json "
                "does not contain usable failed-fund details."
            )
            print(
                "Recovery 2 reconstructed the Recovery 1 "
                "failed-fund list from Recovery 1 failure.json files."
            )
            print(
                "Baseline output_holdings/ was NOT used."
            )

    if not normalized_failed:

        raise RuntimeError(
            "Recovery 1 produced no usable failed-fund details. "
            "Recovery 2 will NOT fall back to the baseline."
        )

    cleaned = []

    seen = set()

    for item in normalized_failed:

        row = item.get(
            "excelRow"
        )

        try:

            row = int(row)

        except Exception as error:

            raise RuntimeError(
                "Recovery 1 failure entry has invalid excelRow: "
                f"{item.get('excelRow')}"
            ) from error

        if row in seen:

            raise RuntimeError(
                "Duplicate Recovery 1 failed Excel row: "
                f"{row}"
            )

        seen.add(row)

        entry = dict(item)

        entry["excelRow"] = row

        cleaned.append(entry)

    cleaned.sort(
        key=lambda item: item["excelRow"]
    )

    payload[
        "failedFundsDetail"
    ] = cleaned

    payload[
        "_recovery2FailureSource"
    ] = source

    return payload


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

            url = clean_text(
                worksheet.cell(
                    row=row_number,
                    column=1,
                ).value
            )

            pruaccess_name = clean_text(
                worksheet.cell(
                    row=row_number,
                    column=2,
                ).value
            )

            if not url:
                continue

            funds[row_number] = {
                "excelRow": row_number,
                "prudentialUrl": url,
                "pruAccessName": pruaccess_name,
            }

    finally:

        workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated Prudential URLs found "
            "in Excel Column A."
        )

    return funds


def build_recovery_universe(
    recovery_1: dict,
    excel_funds: dict[int, dict],
) -> list[dict]:

    failed_entries = recovery_1.get(
        "failedFundsDetail"
    )

    if not isinstance(
        failed_entries,
        list,
    ):

        raise RuntimeError(
            "Recovery 1 failedFundsDetail is invalid."
        )

    universe = []

    seen_rows = set()

    for failure in failed_entries:

        if not isinstance(
            failure,
            dict,
        ):

            raise RuntimeError(
                "Recovery 1 failure entry is not an object."
            )

        try:

            excel_row = int(
                failure["excelRow"]
            )

        except Exception as error:

            raise RuntimeError(
                "Recovery 1 failure entry has invalid "
                f"excelRow: {failure.get('excelRow')}"
            ) from error

        if excel_row in seen_rows:

            raise RuntimeError(
                "Duplicate Recovery 1 failed Excel row: "
                f"{excel_row}"
            )

        seen_rows.add(
            excel_row
        )

        if excel_row not in excel_funds:

            raise RuntimeError(
                f"Recovery 1 failed row {excel_row} "
                "does not exist in Excel."
            )

        excel_fund = excel_funds[
            excel_row
        ]

        excel_url = clean_text(
            excel_fund.get(
                "prudentialUrl"
            )
        )

        excel_name = clean_text(
            excel_fund.get(
                "pruAccessName"
            )
        )

        failure_url = clean_text(
            failure.get(
                "prudentialUrl"
            )
        )

        failure_name = clean_text(
            failure.get(
                "excelPruAccessName"
            )
            or failure.get(
                "pruAccessName"
            )
        )

        if (
            failure_url
            and failure_url != excel_url
        ):

            raise RuntimeError(
                "Recovery 1 URL does not match Excel "
                f"for row {excel_row}."
            )

        if (
            failure_name
            and excel_name
            and failure_name != excel_name
        ):

            raise RuntimeError(
                "Recovery 1 PruAccess name does not match "
                f"Excel for row {excel_row}."
            )

        universe.append(
            {
                "excelRow": excel_row,
                "prudentialUrl": excel_url,
                "pruAccessName": excel_name,
                "recovery1Failure": failure,
            }
        )

    return universe


# =============================================================================
# PRUDENTIAL PAGE / FACTSHEET
# =============================================================================

def find_factsheet_url(
    page,
    prudential_url: str,
) -> str:

    ensure_prudential_url(
        prudential_url
    )

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        POST_PAGE_WAIT_MS
    )

    links = page.locator(
        "a"
    )

    candidates = []

    for index in range(
        links.count()
    ):

        anchor = links.nth(
            index
        )

        try:

            href = clean_text(
                anchor.get_attribute(
                    "href"
                )
            )

            text = clean_text(
                anchor.inner_text(
                    timeout=3000
                )
            )

        except Exception:

            continue

        if not href:
            continue

        absolute = urljoin(
            page.url,
            href,
        )

        if not is_prudential_url(
            absolute
        ):

            continue

        combined = normalize_text(
            f"{text} {absolute}"
        )

        score = 0

        if "factsheet" in combined:
            score += 100

        if "fund factsheet" in combined:
            score += 50

        if ".pdf" in combined:
            score += 50

        if "fund" in combined:
            score += 10

        if score > 0:

            candidates.append(
                (
                    score,
                    absolute,
                )
            )

    if not candidates:

        raise RuntimeError(
            "Could not locate an official Prudential "
            "factsheet link on the fund page."
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

    factsheet_url = ensure_prudential_url(
        factsheet_url
    )

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    if response.status != 200:

        raise RuntimeError(
            "Factsheet download returned "
            f"HTTP {response.status}."
        )

    body = response.body()

    if not body.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Official factsheet response is not a PDF."
        )

    return body


# =============================================================================
# PDF EXTRACTION
# =============================================================================

def read_pdf_pages(
    factsheet_bytes: bytes,
) -> list[dict]:

    reader = PdfReader(
        BytesIO(factsheet_bytes)
    )

    pages = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            text = pdf_page.extract_text() or ""

        except Exception:

            text = ""

        pages.append(
            {
                "page": page_number,
                "text": text,
                "lines": [
                    clean_text(line)
                    for line in text.splitlines()
                    if clean_text(line)
                ],
            }
        )

    return pages


def extract_pdf_text(
    factsheet_bytes: bytes,
) -> tuple[str, int, list[dict]]:

    pages = read_pdf_pages(
        factsheet_bytes
    )

    full_text = "\n".join(
        page["text"]
        for page in pages
    )

    return (
        full_text,
        len(pages),
        pages,
    )


# =============================================================================
# POSITIONED PDF EXTRACTION
# =============================================================================

def positioned_pdf_words(
    factsheet_bytes: bytes,
) -> list[dict]:

    reader = PdfReader(
        BytesIO(factsheet_bytes)
    )

    positioned = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        words = []

        try:

            def visitor_text(
                text,
                cm,
                tm,
                font_dict,
                font_size,
            ):

                if not text:
                    return

                raw_value = str(
                    text
                )

                value = clean_text(
                    raw_value
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

                    size = float(
                        font_size or 0
                    )

                except Exception:

                    size = 0.0

                words.append(
                    {
                        "page": page_number,
                        "text": value,
                        "rawText": raw_value,
                        "x": x,
                        "y": y,
                        "fontSize": size,
                    }
                )

            pdf_page.extract_text(
                visitor_text=visitor_text
            )

        except Exception:

            words = []

        positioned.extend(
            words
        )

    return positioned


def group_positioned_rows(
    positioned: list[dict],
    y_tolerance: float = 3.0,
) -> list[dict]:

    by_page = {}

    for item in positioned:

        by_page.setdefault(
            int(item["page"]),
            [],
        ).append(item)

    output = []

    for page_number in sorted(
        by_page
    ):

        items = sorted(
            by_page[page_number],
            key=lambda item: (
                -float(item.get("y", 0)),
                float(item.get("x", 0)),
            ),
        )

        rows = []

        for item in items:

            y = float(
                item.get(
                    "y",
                    0,
                )
            )

            selected = None

            for row in rows:

                if abs(
                    row["y"] - y
                ) <= y_tolerance:

                    selected = row
                    break

            if selected is None:

                selected = {
                    "page": page_number,
                    "y": y,
                    "items": [],
                }

                rows.append(
                    selected
                )

            selected["items"].append(
                item
            )

        for row in rows:

            row["items"].sort(
                key=lambda item: float(
                    item.get(
                        "x",
                        0,
                    )
                )
            )

            row["text"] = clean_text(
                " ".join(
                    item.get(
                        "text",
                        "",
                    )
                    for item in row["items"]
                )
            )

        rows.sort(
            key=lambda row: -row["y"]
        )

        output.extend(
            rows
        )

    return output


# =============================================================================
# DIAGNOSTIC SEARCH
# =============================================================================

def find_heading_candidates(
    pages: list[dict],
) -> list[dict]:

    matches = []

    for page in pages:

        page_number = page[
            "page"
        ]

        for line_index, line in enumerate(
            page["lines"]
        ):

            normalized = normalize_text(
                line
            )

            for pattern in SECTION_HEADINGS:

                if pattern.search(
                    normalized
                ):

                    matches.append(
                        {
                            "page": page_number,
                            "lineIndex": line_index,
                            "text": line,
                            "normalized": normalized,
                            "type": "section_heading",
                        }
                    )

                    break

    return matches


def find_keyword_candidates(
    pages: list[dict],
) -> list[dict]:

    matches = []

    for page in pages:

        page_number = page[
            "page"
        ]

        for line_index, line in enumerate(
            page["lines"]
        ):

            normalized = normalize_text(
                line
            )

            found = [
                keyword
                for keyword in HOLDINGS_KEYWORDS
                if keyword in normalized
            ]

            if found:

                matches.append(
                    {
                        "page": page_number,
                        "lineIndex": line_index,
                        "text": line,
                        "keywords": found,
                    }
                )

    return matches


def find_percentage_candidates(
    pages: list[dict],
) -> list[dict]:

    matches = []

    for page in pages:

        page_number = page[
            "page"
        ]

        for line_index, line in enumerate(
            page["lines"]
        ):

            percentages = []

            for match in percentage_matches(
                line
            ):

                try:

                    value = float(
                        match.group(1)
                    )

                except Exception:

                    continue

                percentages.append(
                    {
                        "value": value,
                        "text": match.group(0),
                        "start": match.start(),
                        "end": match.end(),
                    }
                )

            if percentages:

                matches.append(
                    {
                        "page": page_number,
                        "lineIndex": line_index,
                        "text": line,
                        "percentages": percentages,
                    }
                )

    return matches


def find_rank_candidates(
    pages: list[dict],
) -> list[dict]:

    matches = []

    for page in pages:

        page_number = page[
            "page"
        ]

        lines = page[
            "lines"
        ]

        for line_index, line in enumerate(
            lines
        ):

            match = LEADING_RANK_RE.match(
                line
            )

            if match:

                rank = int(
                    match.group(1)
                )

                if 1 <= rank <= 10:

                    matches.append(
                        {
                            "page": page_number,
                            "lineIndex": line_index,
                            "rank": rank,
                            "text": line,
                            "remainder": clean_text(
                                match.group(2)
                            ),
                            "hasPercentage": bool(
                                percentage_matches(
                                    line
                                )
                            ),
                        }
                    )

                continue

            rank_only = RANK_ONLY_RE.match(
                line
            )

            if rank_only:

                rank = int(
                    rank_only.group(1)
                )

                if 1 <= rank <= 10:

                    matches.append(
                        {
                            "page": page_number,
                            "lineIndex": line_index,
                            "rank": rank,
                            "text": line,
                            "remainder": "",
                            "hasPercentage": False,
                        }
                    )

    return matches


def surrounding_lines(
    pages: list[dict],
    page_number: int,
    line_index: int,
    radius: int = 4,
) -> list[str]:

    page = next(
        (
            item
            for item in pages
            if item["page"] == page_number
        ),
        None,
    )

    if page is None:
        return []

    start = max(
        0,
        line_index - radius,
    )

    end = min(
        len(page["lines"]),
        line_index + radius + 1,
    )

    return page["lines"][
        start:end
    ]


# =============================================================================
# TEXT LINE HELPERS
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

        if (
            current
            and current.endswith("-")
            and index + 1 < len(lines)
        ):

            next_line = clean_text(
                lines[index + 1]
            )

            if (
                next_line
                and not next_line.startswith("%")
            ):

                current = (
                    current[:-1]
                    + next_line
                )

                index += 1

        if current:

            merged.append(
                current
            )

        index += 1

    return merged


def clean_holding_name(
    value: str,
) -> str:

    name = clean_text(
        value
    )

    name = re.sub(
        r"^[•·▪■□*]+",
        "",
        name,
    ).strip()

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
    ).strip(
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

    text = clean_text(
        value
    )

    text = text.replace(
        "|",
        " ",
    )

    text = re.sub(
        r"\bNone\b",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"^[•·▪■□*]+",
        "",
        text,
    )

    return clean_text(
        text
    )


def combine_holding_name_fragments(
    fragments: list[str],
) -> str:

    cleaned = []

    for fragment in fragments:

        fragment = clean_holding_fragment(
            fragment
        )

        if fragment:
            cleaned.append(
                fragment
            )

    return clean_holding_name(
        " ".join(
            cleaned
        )
    )


def percentage_matches(
    line: str,
):

    return list(
        PERCENTAGE_RE.finditer(
            line
        )
    )


def find_last_percentage_in_line(
    line: str,
):

    matches = percentage_matches(
        line
    )

    if not matches:
        return None

    match = matches[-1]

    return (
        float(
            match.group(1)
        ),
        match.group(0),
        match.start(),
        match.end(),
    )


def extract_leading_rank(
    line: str,
):

    text = clean_text(
        line
    )

    match = LEADING_RANK_RE.match(
        text
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

    return None, text


def is_noise_line(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return True

    if normalized in {
        "top 10 holdings",
        "top ten holdings",
        "top holdings",
        "holding",
        "holdings",
        "weight",
        "weights",
        "percentage",
        "%",
        "name",
        "names",
        "security",
        "securities",
        "investment",
        "investments",
        "portfolio",
    }:

        return True

    return False


# =============================================================================
# HOLDINGS SECTION
# =============================================================================

def extract_explicit_holdings_section(
    pages: list[dict],
) -> tuple[str, str, dict | None]:

    headings = find_heading_candidates(
        pages
    )

    if not headings:

        return (
            "",
            "not_found",
            None,
        )

    heading = headings[0]

    page_number = heading[
        "page"
    ]

    page = next(
        item
        for item in pages
        if item["page"] == page_number
    )

    start = heading[
        "lineIndex"
    ] + 1

    collected = []

    for line in page["lines"][start:]:

        normalized = normalize_text(
            line
        )

        if normalized.startswith(
            "source"
        ):

            break

        if normalized.startswith(
            "inception date"
        ):

            break

        if normalized.startswith(
            "important information"
        ):

            break

        if normalized.startswith(
            "disclaimer"
        ):

            break

        if normalized.startswith(
            "past performance"
        ):

            break

        if normalized.startswith(
            "portfolio characteristics"
        ):

            break

        if normalized.startswith(
            "asset allocation"
        ):

            break

        if not line:
            continue

        collected.append(
            line
        )

    section = "\n".join(
        collected
    ).strip()

    if not section:

        return (
            "",
            "published_empty",
            heading,
        )

    return (
        section,
        "published",
        heading,
    )


# =============================================================================
# STRICT HOLDINGS PARSER
# =============================================================================

def validate_holdings(
    holdings: list[dict],
) -> list[dict]:

    if not holdings:

        raise HoldingsParseFailure(
            "No holding/percentage pairs were extracted."
        )

    if len(holdings) > MAX_HOLDINGS:

        raise HoldingsParseFailure(
            f"Parser produced more than {MAX_HOLDINGS} holdings."
        )

    actual = [
        item.get(
            "rank"
        )
        for item in holdings
    ]

    expected = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    if actual != expected:

        raise HoldingsParseFailure(
            "Holding ranks are not sequential. "
            f"Parsed={actual}; Expected={expected}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding.get(
                "name"
            )
        )

        if not name:

            raise HoldingsParseFailure(
                "Holding name is empty."
            )

        weight = holding.get(
            "weightPercent"
        )

        if not isinstance(
            weight,
            (int, float),
        ):

            raise HoldingsParseFailure(
                "Holding weight is invalid."
            )

        if weight < 0 or weight > 100:

            raise HoldingsParseFailure(
                "Holding weight is outside 0-100%."
            )

        holding["name"] = name

    return holdings


def parse_holdings_primary(
    section_text: str,
) -> list[dict]:

    logical_lines = merge_hyphenated_line_breaks(
        [
            clean_text(line)
            for line in section_text.splitlines()
            if clean_text(line)
        ]
    )

    holdings = []

    pending_fragments = []

    pending_rank = None

    for raw_line in logical_lines:

        line = clean_text(
            raw_line
        )

        if is_noise_line(
            line
        ):

            continue

        rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if rank is not None:

            if pending_fragments:

                raise HoldingsParseFailure(
                    "New holding rank appeared before "
                    "the previous holding received a percentage."
                )

            pending_rank = rank

            line = remainder

        matches = percentage_matches(
            line
        )

        if not matches:

            if (
                pending_rank is not None
                or holdings
            ):

                fragment = clean_holding_fragment(
                    line
                )

                if fragment:

                    pending_fragments.append(
                        fragment
                    )

            continue

        if len(matches) != 1:

            raise HoldingsParseFailure(
                "Primary parser encountered multiple "
                "percentages on one logical line."
            )

        match = matches[0]

        percentage = float(
            match.group(1)
        )

        percentage_text = match.group(0)

        fragment = clean_holding_fragment(
            line[
                :match.start()
            ]
        )

        if fragment:

            pending_fragments.append(
                fragment
            )

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:

            raise HoldingsParseFailure(
                "Published percentage has no holding name."
            )

        holdings.append(
            {
                "rank": (
                    pending_rank
                    if pending_rank is not None
                    else len(holdings) + 1
                ),
                "name": name,
                "weightPercent": percentage,
                "weightText": percentage_text,
            }
        )

        pending_fragments = []

        pending_rank = None

    if pending_fragments:

        raise HoldingsParseFailure(
            "Trailing holding text has no published percentage."
        )

    return validate_holdings(
        holdings
    )


def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:

    logical_lines = merge_hyphenated_line_breaks(
        [
            clean_text(line)
            for line in section_text.splitlines()
            if clean_text(line)
        ]
    )

    holdings = []

    pending_fragments = []

    pending_rank = None

    def commit(
        percentage,
        percentage_text,
    ):

        nonlocal pending_fragments
        nonlocal pending_rank

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:

            raise HoldingsParseFailure(
                "Fallback parser found a published "
                "percentage but no holding name."
            )

        if len(holdings) >= MAX_HOLDINGS:

            raise HoldingsParseFailure(
                "Fallback parser produced more than "
                f"{MAX_HOLDINGS} holdings."
            )

        holdings.append(
            {
                "rank": (
                    pending_rank
                    if pending_rank is not None
                    else len(holdings) + 1
                ),
                "name": name,
                "weightPercent": percentage,
                "weightText": percentage_text,
            }
        )

        pending_fragments = []

        pending_rank = None

    for raw_line in logical_lines:

        line = clean_text(
            raw_line
        )

        if is_noise_line(
            line
        ):

            continue

        rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if rank is not None:

            if pending_fragments:

                raise HoldingsParseFailure(
                    "Fallback parser encountered a new "
                    "rank before the previous holding "
                    "received a published percentage."
                )

            pending_rank = rank

            line = remainder

        last_percentage = (
            find_last_percentage_in_line(
                line
            )
        )

        if last_percentage is None:

            fragment = clean_holding_fragment(
                line
            )

            if fragment:

                pending_fragments.append(
                    fragment
                )

            continue

        percentage, percentage_text, start, end = (
            last_percentage
        )

        fragment = clean_holding_fragment(
            line[
                :start
            ]
        )

        if fragment:

            pending_fragments.append(
                fragment
            )

        commit(
            percentage,
            percentage_text,
        )

    if pending_fragments:

        raise HoldingsParseFailure(
            "Fallback parser found trailing holding text "
            "without a published percentage."
        )

    return validate_holdings(
        holdings
    )


# =============================================================================
# ADAPTIVE CANDIDATE REGIONS
# =============================================================================

def build_text_candidate_regions(
    pages: list[dict],
) -> list[dict]:

    """
    Build candidate text regions without requiring a Top 10 Holdings heading.

    A candidate begins around a rank 1-10 or a holdings keyword and continues
    through nearby lines until a strong document boundary is encountered.

    This is candidate discovery only. Strict validation decides whether a
    candidate is actually accepted.
    """

    regions = []

    for page in pages:

        lines = page[
            "lines"
        ]

        page_number = page[
            "page"
        ]

        rank_indices = []

        for index, line in enumerate(
            lines
        ):

            match = LEADING_RANK_RE.match(
                line
            )

            if match:

                rank = int(
                    match.group(1)
                )

                if 1 <= rank <= 10:

                    rank_indices.append(
                        index
                    )

        keyword_indices = []

        for index, line in enumerate(
            lines
        ):

            normalized = normalize_text(
                line
            )

            if any(
                keyword in normalized
                for keyword in HOLDINGS_KEYWORDS
            ):

                keyword_indices.append(
                    index
                )

        starts = sorted(
            set(
                rank_indices
                + keyword_indices
            )
        )

        for start in starts:

            end = min(
                len(lines),
                start + 40,
            )

            region_lines = []

            for index in range(
                start,
                end,
            ):

                line = lines[index]

                normalized = normalize_text(
                    line
                )

                if index > start:

                    if (
                        normalized.startswith(
                            "important information"
                        )
                        or normalized.startswith(
                            "disclaimer"
                        )
                        or normalized.startswith(
                            "past performance"
                        )
                        or normalized.startswith(
                            "source"
                        )
                    ):

                        break

                region_lines.append(
                    line
                )

            if not region_lines:
                continue

            text = "\n".join(
                region_lines
            )

            rank_count = len(
                [
                    line
                    for line in region_lines
                    if (
                        LEADING_RANK_RE.match(
                            line
                        )
                        and 1 <= int(
                            LEADING_RANK_RE.match(
                                line
                            ).group(1)
                        ) <= 10
                    )
                ]
            )

            percentage_count = sum(
                len(
                    percentage_matches(
                        line
                    )
                )
                for line in region_lines
            )

            if (
                rank_count == 0
                and percentage_count == 0
            ):

                continue

            regions.append(
                {
                    "page": page_number,
                    "startLineIndex": start,
                    "endLineIndex": (
                        start
                        + len(region_lines)
                        - 1
                    ),
                    "text": text,
                    "rankCount": rank_count,
                    "percentageCount": percentage_count,
                }
            )

    # Deduplicate identical candidate text.
    deduplicated = []

    seen = set()

    for region in regions:

        key = (
            region["page"],
            region["startLineIndex"],
            region["endLineIndex"],
            region["text"],
        )

        if key in seen:
            continue

        seen.add(key)

        deduplicated.append(
            region
        )

    # Highest evidence first.
    deduplicated.sort(
        key=lambda item: (
            -item["rankCount"],
            -item["percentageCount"],
            item["page"],
            item["startLineIndex"],
        )
    )

    return deduplicated


def extract_ranked_text_windows(
    pages: list[dict],
) -> list[dict]:

    """
    More conservative windows.

    Each candidate starts at a detected rank and includes a bounded number
    of subsequent lines. This is useful when PDF text extraction separates
    the security name and percentage vertically.
    """

    candidates = []

    for page in pages:

        lines = page[
            "lines"
        ]

        page_number = page[
            "page"
        ]

        for start, line in enumerate(
            lines
        ):

            match = LEADING_RANK_RE.match(
                line
            )

            if not match:
                continue

            rank = int(
                match.group(1)
            )

            if not 1 <= rank <= 10:
                continue

            for width in (
                8,
                12,
                18,
                25,
            ):

                end = min(
                    len(lines),
                    start + width,
                )

                region_lines = lines[
                    start:end
                ]

                text = "\n".join(
                    region_lines
                )

                ranks = []

                for candidate_line in region_lines:

                    candidate_match = (
                        LEADING_RANK_RE.match(
                            candidate_line
                        )
                    )

                    if candidate_match:

                        candidate_rank = int(
                            candidate_match.group(
                                1
                            )
                        )

                        if 1 <= candidate_rank <= 10:

                            ranks.append(
                                candidate_rank
                            )

                percentages = sum(
                    len(
                        percentage_matches(
                            candidate_line
                        )
                    )
                    for candidate_line in region_lines
                )

                if percentages == 0:
                    continue

                candidates.append(
                    {
                        "page": page_number,
                        "startLineIndex": start,
                        "endLineIndex": end - 1,
                        "text": text,
                        "rankCount": len(ranks),
                        "ranks": ranks,
                        "percentageCount": percentages,
                        "windowWidth": width,
                    }
                )

    candidates.sort(
        key=lambda item: (
            -item["rankCount"],
            -item["percentageCount"],
            item["page"],
            item["startLineIndex"],
            item["windowWidth"],
        )
    )

    return candidates


# =============================================================================
# SPATIAL DIAGNOSTICS
# =============================================================================

def spatial_keyword_matches(
    positioned_rows: list[dict],
) -> list[dict]:

    matches = []

    for row in positioned_rows:

        normalized = normalize_text(
            row.get(
                "text",
                ""
            )
        )

        found = [
            keyword
            for keyword in HOLDINGS_KEYWORDS
            if keyword in normalized
        ]

        if found:

            matches.append(
                {
                    "page": row.get(
                        "page"
                    ),
                    "y": row.get(
                        "y"
                    ),
                    "text": row.get(
                        "text"
                    ),
                    "keywords": found,
                    "items": row.get(
                        "items",
                        [],
                    ),
                }
            )

    return matches


def spatial_rank_matches(
    positioned_rows: list[dict],
) -> list[dict]:

    matches = []

    for row in positioned_rows:

        text = clean_text(
            row.get(
                "text",
                ""
            )
        )

        match = LEADING_RANK_RE.match(
            text
        )

        if match:

            rank = int(
                match.group(
                    1
                )
            )

            if 1 <= rank <= 10:

                matches.append(
                    {
                        "page": row.get(
                            "page"
                        ),
                        "y": row.get(
                            "y"
                        ),
                        "rank": rank,
                        "text": text,
                        "items": row.get(
                            "items",
                            [],
                        ),
                    }
                )

        rank_only = RANK_ONLY_RE.match(
            text
        )

        if rank_only:

            rank = int(
                rank_only.group(
                    1
                )
            )

            if 1 <= rank <= 10:

                matches.append(
                    {
                        "page": row.get(
                            "page"
                        ),
                        "y": row.get(
                            "y"
                        ),
                        "rank": rank,
                        "text": text,
                        "items": row.get(
                            "items",
                            [],
                        ),
                    }
                )

    return matches


def spatial_percentage_matches(
    positioned_rows: list[dict],
) -> list[dict]:

    matches = []

    for row in positioned_rows:

        text = clean_text(
            row.get(
                "text",
                ""
            )
        )

        percentages = []

        for match in percentage_matches(
            text
        ):

            percentages.append(
                {
                    "value": float(
                        match.group(1)
                    ),
                    "text": match.group(0),
                    "start": match.start(),
                    "end": match.end(),
                }
            )

        if percentages:

            matches.append(
                {
                    "page": row.get(
                        "page"
                    ),
                    "y": row.get(
                        "y"
                    ),
                    "text": text,
                    "percentages": percentages,
                    "items": row.get(
                        "items",
                        [],
                    ),
                }
            )

    return matches


def build_spatial_candidate_windows(
    positioned_rows: list[dict],
) -> list[dict]:

    """
    Build spatial windows from actual PDF rows.

    This does NOT pair names to percentages merely by distance.

    Instead, each candidate window is retained as raw evidence. The strict
    parser is allowed to accept it only when the resulting sequence itself
    contains explicit rank/name/percentage evidence.
    """

    candidates = []

    by_page = {}

    for row in positioned_rows:

        by_page.setdefault(
            int(row["page"]),
            [],
        ).append(row)

    for page_number in sorted(
        by_page
    ):

        rows = by_page[
            page_number
        ]

        rows.sort(
            key=lambda row: -float(
                row.get(
                    "y",
                    0,
                )
            )
        )

        for start_index, row in enumerate(
            rows
        ):

            text = clean_text(
                row.get(
                    "text",
                    ""
                )
            )

            rank_match = LEADING_RANK_RE.match(
                text
            )

            rank_only_match = RANK_ONLY_RE.match(
                text
            )

            rank = None

            if rank_match:

                rank = int(
                    rank_match.group(
                        1
                    )
                )

            elif rank_only_match:

                rank = int(
                    rank_only_match.group(
                        1
                    )
                )

            if rank is None or not 1 <= rank <= 10:
                continue

            for width in (
                6,
                10,
                15,
                20,
            ):

                end = min(
                    len(rows),
                    start_index + width,
                )

                window = rows[
                    start_index:end
                ]

                percentage_rows = [
                    candidate
                    for candidate in window
                    if percentage_matches(
                        candidate.get(
                            "text",
                            ""
                        )
                    )
                ]

                if not percentage_rows:
                    continue

                candidate = {
                    "page": page_number,
                    "startRowIndex": start_index,
                    "endRowIndex": end - 1,
                    "rank": rank,
                    "rowCount": len(window),
                    "percentageRowCount": len(
                        percentage_rows
                    ),
                    "rows": window,
                }

                candidates.append(
                    candidate
                )

    candidates.sort(
        key=lambda item: (
            -item["percentageRowCount"],
            item["page"],
            item["startRowIndex"],
            item["rowCount"],
        )
    )

    return candidates


# =============================================================================
# SPATIAL TABLE RECONSTRUCTION
# =============================================================================

def spatial_window_to_lines(
    candidate: dict,
) -> list[str]:

    lines = []

    for row in candidate.get(
        "rows",
        [],
    ):

        text = clean_text(
            row.get(
                "text",
                ""
            )
        )

        if text:
            lines.append(
                text
            )

    return lines


def spatial_candidate_to_json(
    candidate: dict,
) -> dict:

    rows = []

    for row in candidate.get(
        "rows",
        [],
    ):

        rows.append(
            {
                "page": row.get(
                    "page"
                ),
                "y": row.get(
                    "y"
                ),
                "text": row.get(
                    "text"
                ),
                "items": [
                    {
                        "text": item.get(
                            "text"
                        ),
                        "x": item.get(
                            "x"
                        ),
                        "y": item.get(
                            "y"
                        ),
                        "fontSize": item.get(
                            "fontSize"
                        ),
                    }
                    for item in row.get(
                        "items",
                        [],
                    )
                ],
            }
        )

    return {
        "page": candidate.get(
            "page"
        ),
        "startRowIndex": candidate.get(
            "startRowIndex"
        ),
        "endRowIndex": candidate.get(
            "endRowIndex"
        ),
        "rank": candidate.get(
            "rank"
        ),
        "rowCount": candidate.get(
            "rowCount"
        ),
        "percentageRowCount": candidate.get(
            "percentageRowCount"
        ),
        "rows": rows,
    }


# =============================================================================
# ADAPTIVE EXTRACTION
# =============================================================================

def attempt_explicit_section(
    pages: list[dict],
) -> dict:

    section_text, status, heading = (
        extract_explicit_holdings_section(
            pages
        )
    )

    diagnostic = {
        "status": status,
        "heading": heading,
    }

    if status != "published":

        return {
            "success": False,
            "method": "explicit_section",
            "error": (
                "No explicit published holdings section."
            ),
            "diagnostic": diagnostic,
        }

    primary_error = ""

    try:

        holdings = parse_holdings_primary(
            section_text
        )

        return {
            "success": True,
            "method": "explicit_section_primary",
            "holdings": holdings,
            "sectionText": section_text,
            "error": "",
            "diagnostic": diagnostic,
        }

    except Exception as error:

        primary_error = clean_text(
            str(error)
        )

    try:

        holdings = parse_holdings_fallback(
            section_text
        )

        return {
            "success": True,
            "method": "explicit_section_fallback",
            "holdings": holdings,
            "sectionText": section_text,
            "error": primary_error,
            "diagnostic": diagnostic,
        }

    except Exception as error:

        return {
            "success": False,
            "method": "explicit_section",
            "error": (
                "Primary: "
                f"{primary_error}; "
                "Fallback: "
                f"{clean_text(str(error))}"
            ),
            "diagnostic": diagnostic,
        }


def attempt_text_candidates(
    pages: list[dict],
) -> list[dict]:

    results = []

    regions = build_text_candidate_regions(
        pages
    )

    ranked_windows = (
        extract_ranked_text_windows(
            pages
        )
    )

    candidates = (
        regions
        + ranked_windows
    )

    for index, candidate in enumerate(
        candidates,
        start=1,
    ):

        text = candidate.get(
            "text",
            ""
        )

        if not text:
            continue

        attempts = []

        try:

            holdings = parse_holdings_primary(
                text
            )

            attempts.append(
                {
                    "success": True,
                    "parser": "primary",
                    "holdings": holdings,
                }
            )

        except Exception as error:

            attempts.append(
                {
                    "success": False,
                    "parser": "primary",
                    "error": clean_text(
                        str(error)
                    ),
                }
            )

        try:

            holdings = parse_holdings_fallback(
                text
            )

            attempts.append(
                {
                    "success": True,
                    "parser": "fallback",
                    "holdings": holdings,
                }
            )

        except Exception as error:

            attempts.append(
                {
                    "success": False,
                    "parser": "fallback",
                    "error": clean_text(
                        str(error)
                    ),
                }
            )

        for attempt in attempts:

            if attempt.get(
                "success"
            ):

                results.append(
                    {
                        "candidateIndex": index,
                        "candidate": candidate,
                        "parser": attempt[
                            "parser"
                        ],
                        "holdings": attempt[
                            "holdings"
                        ],
                    }
                )

    return results


def attempt_spatial_candidates(
    positioned_rows: list[dict],
) -> list[dict]:

    candidates = (
        build_spatial_candidate_windows(
            positioned_rows
        )
    )

    results = []

    for index, candidate in enumerate(
        candidates,
        start=1,
    ):

        lines = spatial_window_to_lines(
            candidate
        )

        text = "\n".join(
            lines
        )

        if not text:
            continue

        for parser_name, parser in (
            (
                "spatial_primary",
                parse_holdings_primary,
            ),
            (
                "spatial_fallback",
                parse_holdings_fallback,
            ),
        ):

            try:

                holdings = parser(
                    text
                )

            except Exception:

                continue

            results.append(
                {
                    "candidateIndex": index,
                    "candidate": candidate,
                    "parser": parser_name,
                    "holdings": holdings,
                    "sectionText": text,
                }
            )

    return results


# =============================================================================
# RESULT CONSISTENCY
# =============================================================================

def holding_signature(
    holdings: list[dict],
) -> list[tuple]:

    return [
        (
            int(
                item.get(
                    "rank"
                )
            ),
            normalize_text(
                item.get(
                    "name"
                )
            ),
            float(
                item.get(
                    "weightPercent"
                )
            ),
        )
        for item in holdings
    ]


def choose_adaptive_result(
    explicit_result: dict,
    text_results: list[dict],
    spatial_results: list[dict],
) -> dict | None:

    """
    Select only a result that passes strict structural validation.

    Preference:

    1. Explicit section
    2. Text candidate
    3. Spatial candidate

    Multiple independently discovered candidates with different signatures
    are NOT silently reconciled. That ambiguity is retained as diagnostics.
    """

    if explicit_result.get(
        "success"
    ):

        return {
            "source": "explicit",
            "method": explicit_result.get(
                "method"
            ),
            "holdings": explicit_result.get(
                "holdings"
            ),
            "sectionText": explicit_result.get(
                "sectionText"
            ),
            "candidate": explicit_result.get(
                "diagnostic"
            ),
        }

    successful = (
        text_results
        + spatial_results
    )

    if not successful:
        return None

    signatures = {}

    for item in successful:

        signature = tuple(
            holding_signature(
                item["holdings"]
            )
        )

        signatures.setdefault(
            signature,
            [],
        ).append(
            item
        )

    # If only one distinct signature was discovered, it is structurally
    # consistent. If multiple signatures exist, do not guess.
    if len(signatures) != 1:

        return None

    signature = next(
        iter(signatures)
    )

    matching = signatures[
        signature
    ]

    first = matching[0]

    return {
        "source": (
            "text_candidate"
            if first.get(
                "parser",
                ""
            ).startswith(
                "primary"
            )
            or first.get(
                "parser",
                ""
            ).startswith(
                "fallback"
            )
            else "spatial_candidate"
        ),
        "method": first.get(
            "parser"
        ),
        "holdings": first.get(
            "holdings"
        ),
        "sectionText": first.get(
            "sectionText"
        )
        or first.get(
            "candidate",
            {},
        ).get(
            "text",
            ""
        ),
        "candidate": first.get(
            "candidate"
        ),
        "candidateCountWithSameSignature": len(
            matching
        ),
    }


# =============================================================================
# FORENSIC DIAGNOSTICS
# =============================================================================

def build_page_diagnostics(
    pages: list[dict],
) -> list[dict]:

    output = []

    for page in pages:

        page_number = page[
            "page"
        ]

        lines = page[
            "lines"
        ]

        rank_hits = []

        percentage_hits = []

        keyword_hits = []

        for line_index, line in enumerate(
            lines
        ):

            rank_match = LEADING_RANK_RE.match(
                line
            )

            if rank_match:

                rank = int(
                    rank_match.group(
                        1
                    )
                )

                if 1 <= rank <= 10:

                    rank_hits.append(
                        {
                            "lineIndex": line_index,
                            "rank": rank,
                            "text": line,
                            "hasPercentage": bool(
                                percentage_matches(
                                    line
                                )
                            ),
                        }
                    )

            percentages = percentage_matches(
                line
            )

            if percentages:

                percentage_hits.append(
                    {
                        "lineIndex": line_index,
                        "text": line,
                        "percentages": [
                            {
                                "value": float(
                                    match.group(
                                        1
                                    )
                                ),
                                "text": match.group(
                                    0
                                ),
                            }
                            for match in percentages
                        ],
                    }
                )

            normalized = normalize_text(
                line
            )

            found_keywords = [
                keyword
                for keyword in HOLDINGS_KEYWORDS
                if keyword in normalized
            ]

            if found_keywords:

                keyword_hits.append(
                    {
                        "lineIndex": line_index,
                        "text": line,
                        "keywords": found_keywords,
                    }
                )

        output.append(
            {
                "page": page_number,
                "lineCount": len(
                    lines
                ),
                "rankCandidates": rank_hits,
                "percentageCandidates": percentage_hits,
                "keywordCandidates": keyword_hits,
            }
        )

    return output


def build_context_diagnostics(
    pages: list[dict],
    rank_candidates: list[dict],
    percentage_candidates: list[dict],
    keyword_candidates: list[dict],
) -> dict:

    rank_context = []

    for item in rank_candidates:

        rank_context.append(
            {
                "page": item[
                    "page"
                ],
                "lineIndex": item[
                    "lineIndex"
                ],
                "rank": item[
                    "rank"
                ],
                "text": item[
                    "text"
                ],
                "surroundingLines": surrounding_lines(
                    pages,
                    item[
                        "page"
                    ],
                    item[
                        "lineIndex"
                    ],
                    radius=5,
                ),
            }
        )

    percentage_context = []

    for item in percentage_candidates:

        percentage_context.append(
            {
                "page": item[
                    "page"
                ],
                "lineIndex": item[
                    "lineIndex"
                ],
                "text": item[
                    "text"
                ],
                "percentages": item[
                    "percentages"
                ],
                "surroundingLines": surrounding_lines(
                    pages,
                    item[
                        "page"
                    ],
                    item[
                        "lineIndex"
                    ],
                    radius=3,
                ),
            }
        )

    keyword_context = []

    for item in keyword_candidates:

        keyword_context.append(
            {
                "page": item[
                    "page"
                ],
                "lineIndex": item[
                    "lineIndex"
                ],
                "text": item[
                    "text"
                ],
                "keywords": item[
                    "keywords"
                ],
                "surroundingLines": surrounding_lines(
                    pages,
                    item[
                        "page"
                    ],
                    item[
                        "lineIndex"
                    ],
                    radius=3,
                ),
            }
        )

    return {
        "rankContexts": rank_context,
        "percentageContexts": percentage_context,
        "keywordContexts": keyword_context,
    }


def save_forensic_diagnostics(
    directory: Path,
    factsheet_bytes: bytes,
    full_text: str,
    pages: list[dict],
    positioned: list[dict],
    diagnostic: dict,
) -> None:

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        directory / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    (
        directory / "factsheet_text.txt"
    ).write_text(
        full_text,
        encoding="utf-8",
    )

    pages_dir = (
        directory / "pages"
    )

    pages_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    for page in pages:

        page_path = (
            pages_dir
            / (
                f"page_"
                f"{int(page['page']):03d}.txt"
            )
        )

        page_path.write_text(
            page.get(
                "text",
                ""
            ),
            encoding="utf-8",
        )

    save_json(
        directory / "positioned_text.json",
        positioned,
    )

    save_json(
        directory / "diagnostic.json",
        diagnostic,
    )


# =============================================================================
# METADATA
# =============================================================================

def extract_data_as_at(
    text: str,
) -> str:

    for pattern in (
        r"data\s+as\s+at\s*[:\-]?\s*([^\n]+)",
        r"as\s+at\s*[:\-]?\s*([^\n]+)",
    ):

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(
                    1
                )
            )

    return ""


def extract_document_date(
    text: str,
) -> str:

    for pattern in (
        r"document\s+date\s*[:\-]?\s*([^\n]+)",
        r"dated\s*[:\-]?\s*([^\n]+)",
    ):

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(
                    1
                )
            )

    return ""


def extract_fund_page_name(
    page,
) -> str:

    try:

        title = clean_text(
            page.title()
        )

        if title:
            return title

    except Exception:
        pass

    return ""


# =============================================================================
# FUND PROCESSING
# =============================================================================

def fund_output_directory(
    excel_fund: dict,
    page_title: str = "",
) -> Path:

    row = int(
        excel_fund[
            "excelRow"
        ]
    )

    identifier = (
        safe_filename(
            page_title
        )
        or safe_filename(
            excel_fund.get(
                "pruAccessName",
                ""
            )
        )
        or f"fund_{row}"
    )

    return (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{row}_{identifier}"
    )


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

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

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
            "Prudential page redirected to a non-Prudential "
            f"host: {final_url}"
        )

    page_title = extract_fund_page_name(
        page
    )

    factsheet_url = find_factsheet_url(
        page,
        final_url,
    )

    factsheet_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    full_text, pdf_page_count, pages = (
        extract_pdf_text(
            factsheet_bytes
        )
    )

    positioned = positioned_pdf_words(
        factsheet_bytes
    )

    positioned_rows = group_positioned_rows(
        positioned
    )

    headings = find_heading_candidates(
        pages
    )

    keywords = find_keyword_candidates(
        pages
    )

    percentages = find_percentage_candidates(
        pages
    )

    ranks = find_rank_candidates(
        pages
    )

    page_diagnostics = (
        build_page_diagnostics(
            pages
        )
    )

    contexts = build_context_diagnostics(
        pages,
        ranks,
        percentages,
        keywords,
    )

    spatial_keywords = (
        spatial_keyword_matches(
            positioned_rows
        )
    )

    spatial_ranks = (
        spatial_rank_matches(
            positioned_rows
        )
    )

    spatial_percentages = (
        spatial_percentage_matches(
            positioned_rows
        )
    )

    spatial_candidates = (
        build_spatial_candidate_windows(
            positioned_rows
        )
    )

    explicit_result = (
        attempt_explicit_section(
            pages
        )
    )

    text_results = (
        attempt_text_candidates(
            pages
        )
    )

    spatial_results = (
        attempt_spatial_candidates(
            positioned_rows
        )
    )

    adaptive_result = (
        choose_adaptive_result(
            explicit_result,
            text_results,
            spatial_results,
        )
    )

    output_directory = fund_output_directory(
        excel_fund,
        page_title,
    )

    diagnostic = {
        "status": (
            "success"
            if adaptive_result
            else "diagnostic_no_confirmed_result"
        ),
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pageTitle": page_title,
        "excelPruAccessName": excel_fund.get(
            "pruAccessName"
        ),
        "factsheetUrl": factsheet_url,
        "factsheetPageCount": pdf_page_count,
        "factsheetDocumentDate": extract_document_date(
            full_text
        ),
        "factsheetDataAsAt": extract_data_as_at(
            full_text
        ),
        "explicitSectionHeadings": headings,
        "keywordCandidateCount": len(
            keywords
        ),
        "percentageCandidateCount": len(
            percentages
        ),
        "rankCandidateCount": len(
            ranks
        ),
        "keywordCandidates": keywords,
        "percentageCandidates": percentages,
        "rankCandidates": ranks,
        "context": contexts,
        "pageDiagnostics": page_diagnostics,
        "spatialKeywordCandidateCount": len(
            spatial_keywords
        ),
        "spatialRankCandidateCount": len(
            spatial_ranks
        ),
        "spatialPercentageCandidateCount": len(
            spatial_percentages
        ),
        "spatialKeywordCandidates": spatial_keywords,
        "spatialRankCandidates": spatial_ranks,
        "spatialPercentageCandidates": spatial_percentages,
        "spatialCandidateCount": len(
            spatial_candidates
        ),
        "spatialCandidatePreviews": [
            spatial_candidate_to_json(
                candidate
            )
            for candidate in spatial_candidates[
                :50
            ]
        ],
        "explicitExtraction": {
            "success": explicit_result.get(
                "success"
            ),
            "method": explicit_result.get(
                "method"
            ),
            "error": explicit_result.get(
                "error"
            ),
            "diagnostic": explicit_result.get(
                "diagnostic"
            ),
        },
        "textCandidateResultCount": len(
            text_results
        ),
        "textCandidateResults": [
            {
                "candidateIndex": item.get(
                    "candidateIndex"
                ),
                "parser": item.get(
                    "parser"
                ),
                "candidate": item.get(
                    "candidate"
                ),
                "holdings": item.get(
                    "holdings"
                ),
            }
            for item in text_results[
                :50
            ]
        ],
        "spatialExtractionResultCount": len(
            spatial_results
        ),
        "spatialExtractionResults": [
            {
                "candidateIndex": item.get(
                    "candidateIndex"
                ),
                "parser": item.get(
                    "parser"
                ),
                "candidate": spatial_candidate_to_json(
                    item.get(
                        "candidate",
                        {}
                    )
                ),
                "holdings": item.get(
                    "holdings"
                ),
            }
            for item in spatial_results[
                :50
            ]
        ],
        "adaptiveResult": (
            {
                "source": adaptive_result.get(
                    "source"
                ),
                "method": adaptive_result.get(
                    "method"
                ),
                "holdings": adaptive_result.get(
                    "holdings"
                ),
                "candidateCountWithSameSignature":
                    adaptive_result.get(
                        "candidateCountWithSameSignature"
                    ),
            }
            if adaptive_result
            else None
        ),
        "rules": {
            "recovery1FailedFundsOnly": True,
            "noBaselineFallback": True,
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
            "exactFinalPdfVerification": True,
            "exactRankNameWeightSignatureRequired": True,
        },
        "generatedAtUtc": utc_now_iso(),
    }

    save_forensic_diagnostics(
        output_directory,
        factsheet_bytes,
        full_text,
        pages,
        positioned,
        diagnostic,
    )

    if not adaptive_result:

        return {
            "success": False,
            "status": (
                "DIAGNOSTIC_NO_CONFIRMED_RESULT"
            ),
            "diagnostic": diagnostic,
            "factsheetBytes": factsheet_bytes,
            "fullText": full_text,
            "factsheetUrl": factsheet_url,
            "outputDirectory": output_directory,
        }

    holdings = validate_holdings(
        adaptive_result[
            "holdings"
        ]
    )

    result = {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pageTitle": page_title,
        "fundName": page_title,
        "excelPruAccessName": excel_fund.get(
            "pruAccessName"
        ),
        "factsheetUrl": factsheet_url,
        "factsheetDocumentDate": extract_document_date(
            full_text
        ),
        "factsheetDataAsAt": extract_data_as_at(
            full_text
        ),
        "factsheetPageCount": pdf_page_count,
        "topHoldingsCount": len(
            holdings
        ),
        "topHoldings": holdings,
        "holdingsParser": adaptive_result.get(
            "method"
        ),
        "holdingsExtractionSource": adaptive_result.get(
            "source"
        ),
        "candidateEvidenceCount": adaptive_result.get(
            "candidateCountWithSameSignature"
        ),
        "recovery1Failure": excel_fund.get(
            "recovery1Failure"
        ),
        "rules": diagnostic[
            "rules"
        ],
    }

    section_text = clean_text(
        adaptive_result.get(
            "sectionText",
            ""
        )
    )

    return {
        "success": True,
        "status": "success",
        "result": result,
        "factsheetBytes": factsheet_bytes,
        "fullText": full_text,
        "sectionText": section_text,
        "diagnostic": diagnostic,
        "outputDirectory": output_directory,
    }


# =============================================================================
# FINAL EXACT VERIFICATION
# =============================================================================

def verify_exact_against_official_pdf(
    page,
    result: dict,
) -> tuple[
    bytes,
    str,
    str,
    list[dict],
]:

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
            "Factsheet re-download returned "
            f"HTTP {response.status}."
        )

    factsheet_bytes = response.body()

    if not factsheet_bytes.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Factsheet re-download did not return a valid PDF."
        )

    full_text, page_count, pages = (
        extract_pdf_text(
            factsheet_bytes
        )
    )

    positioned = positioned_pdf_words(
        factsheet_bytes
    )

    positioned_rows = group_positioned_rows(
        positioned
    )

    explicit_result = (
        attempt_explicit_section(
            pages
        )
    )

    text_results = (
        attempt_text_candidates(
            pages
        )
    )

    spatial_results = (
        attempt_spatial_candidates(
            positioned_rows
        )
    )

    adaptive_result = (
        choose_adaptive_result(
            explicit_result,
            text_results,
            spatial_results,
        )
    )

    if not adaptive_result:

        raise RuntimeError(
            "No structurally confirmed holding result "
            "was found during final verification."
        )

    verified_holdings = validate_holdings(
        adaptive_result[
            "holdings"
        ]
    )

    original_holdings = (
        result.get(
            "topHoldings"
        )
        or []
    )

    if len(
        verified_holdings
    ) != len(
        original_holdings
    ):

        raise RuntimeError(
            "Holding count changed during final verification: "
            f"initial={len(original_holdings)}, "
            f"verified={len(verified_holdings)}."
        )

    if holding_signature(
        verified_holdings
    ) != holding_signature(
        original_holdings
    ):

        raise RuntimeError(
            "Holding signature changed during final verification. "
            "No Recovery 2 result was accepted."
        )

    result[
        "topHoldings"
    ] = verified_holdings

    result[
        "topHoldingsCount"
    ] = len(
        verified_holdings
    )

    result[
        "finalVerification"
    ] = {
        "status": "verified",
        "verifiedAtUtc": utc_now_iso(),
        "exactSignatureMatch": True,
        "verifiedHoldingCount": len(
            verified_holdings
        ),
        "verificationPageCount": page_count,
        "verificationExtractionMethod":
            adaptive_result.get(
                "method"
            ),
        "verificationExtractionSource":
            adaptive_result.get(
                "source"
            ),
    }

    return (
        factsheet_bytes,
        full_text,
        clean_text(
            adaptive_result.get(
                "sectionText",
                ""
            )
        ),
        verified_holdings,
    )


# =============================================================================
# OUTPUT
# =============================================================================

def save_success(
    result: dict,
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
    diagnostic: dict,
) -> Path:

    directory = fund_output_directory(
        {
            "excelRow": result[
                "excelRow"
            ],
            "pruAccessName": result.get(
                "excelPruAccessName"
            ),
        },
        result.get(
            "fundName"
        ),
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        directory / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    (
        directory / "factsheet_text.txt"
    ).write_text(
        full_text,
        encoding="utf-8",
    )

    (
        directory / "top_holdings_section.txt"
    ).write_text(
        section_text,
        encoding="utf-8",
    )

    save_json(
        directory / "top_holdings.json",
        result,
    )

    save_json(
        directory / "recovery_diagnostics.json",
        diagnostic,
    )

    return directory


def save_unconfirmed_diagnostics(
    excel_fund: dict,
    diagnostic: dict,
) -> Path:

    row = int(
        excel_fund[
            "excelRow"
        ]
    )

    existing_directory = (
        diagnostic.get(
            "_outputDirectory"
        )
    )

    if existing_directory:

        directory = Path(
            existing_directory
        )

    else:

        directory = fund_output_directory(
            excel_fund,
            "",
        )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory / "diagnostic_status.json",
        {
            "status": (
                "DIAGNOSTIC_NO_CONFIRMED_RESULT"
            ),
            "excelRow": row,
            "prudentialUrl": excel_fund[
                "prudentialUrl"
            ],
            "excelPruAccessName": excel_fund.get(
                "pruAccessName"
            ),
            "recovery1Failure": excel_fund.get(
                "recovery1Failure"
            ),
            "message": (
                "The official factsheet was downloaded and "
                "forensically inspected, but no single "
                "structurally consistent holdings signature "
                "was confirmed. No holding data was fabricated "
                "or accepted."
            ),
            "generatedAtUtc": utc_now_iso(),
        },
    )

    return directory


def save_failure(
    excel_fund: dict,
    error_text: str,
) -> Path:

    row = int(
        excel_fund[
            "excelRow"
        ]
    )

    directory = (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{row}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory / "failure.json",
        {
            "status": "failed",
            "excelRow": row,
            "prudentialUrl": excel_fund[
                "prudentialUrl"
            ],
            "excelPruAccessName": excel_fund.get(
                "pruAccessName"
            ),
            "error": clean_text(
                error_text
            ),
            "recoveryStage": "Recovery 2",
            "recovery1Failure": excel_fund.get(
                "recovery1Failure"
            ),
            "failedAtUtc": utc_now_iso(),
        },
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

    print()
    print(
        "################################################################"
    )
    print(
        "VGRAT FMS - PRUDENTIAL TOP HOLDINGS RECOVERY 2"
    )
    print(
        "ADAPTIVE / FORENSIC EXTRACTOR"
    )
    print(
        "################################################################"
    )
    print(
        f"Started UTC: {started_at}"
    )
    print(
        f"Recovery 1 summary: "
        f"{RECOVERY_1_RUN_SUMMARY_FILE}"
    )
    print(
        f"Recovery 2 output:  "
        f"{RECOVERY_OUTPUT_DIR}"
    )
    print()
    print(
        "Recovery source: Recovery 1 failed-fund details ONLY"
    )

    # -------------------------------------------------------------------------
    # LOAD RECOVERY 1
    # -------------------------------------------------------------------------

    recovery_1 = (
        load_recovery_1_run_summary()
    )

    excel_funds = read_excel_funds()

    recovery_universe = (
        build_recovery_universe(
            recovery_1,
            excel_funds,
        )
    )

    recovery1_failed_count = len(
        recovery_1.get(
            "failedFundsDetail",
            [],
        )
    )

    print()
    print(
        "Recovery 1 failure source:"
    )
    print(
        f"  {recovery_1.get('_recovery2FailureSource', 'unknown')}"
    )

    print(
        f"Recovery 1 failed funds: "
        f"{recovery1_failed_count}"
    )

    print(
        "Recovery 1 failed rows:"
    )

    print(
        "  "
        + ", ".join(
            str(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_universe
        )
    )

    print(
        f"Recovery 2 universe: "
        f"{len(recovery_universe)}"
    )

    if not recovery_universe:

        completed_at = utc_now_iso()

        save_json(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        save_json(
            RECOVERY_RUN_SUMMARY_FILE,
            {
                "status": "nothing_to_recover",
                "startedAtUtc": started_at,
                "completedAtUtc": completed_at,
                "excelFile": str(
                    EXCEL_FILE
                ),
                "recovery1RunSummary": str(
                    RECOVERY_1_RUN_SUMMARY_FILE
                ),
                "recovery1FailureSource": recovery_1.get(
                    "_recovery2FailureSource"
                ),
                "recovery1FailedFunds":
                    recovery1_failed_count,
                "recovery2Universe": 0,
                "successfulFunds": 0,
                "diagnosticFunds": 0,
                "failedFunds": 0,
                "totalPublishedTopHoldings": 0,
                "successfulFundsDetail": [],
                "diagnosticFundsDetail": [],
                "failedFundsDetail": [],
            },
        )

        return 0

    # -------------------------------------------------------------------------
    # PROCESS
    # -------------------------------------------------------------------------

    successful = []

    diagnostic_funds = []

    failed = []

    all_holdings = []

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

            total = len(
                recovery_universe
            )

            for index, excel_fund in enumerate(
                recovery_universe,
                start=1,
            ):

                excel_row = int(
                    excel_fund[
                        "excelRow"
                    ]
                )

                print()
                print(
                    "=" * 78
                )
                print(
                    f"RECOVERY 2 FUND "
                    f"{index}/{total}"
                )
                print(
                    f"Excel row: {excel_row}"
                )
                print(
                    f"URL: "
                    f"{excel_fund['prudentialUrl']}"
                )
                print(
                    "=" * 78
                )

                last_error = None

                for attempt in range(
                    1,
                    RETRY_COUNT + 1,
                ):

                    try:

                        print(
                            f"Attempt "
                            f"{attempt}/{RETRY_COUNT}"
                        )

                        payload = recover_single_fund(
                            page,
                            excel_fund,
                        )

                        # -----------------------------------------------------
                        # Diagnostic-only outcome
                        # -----------------------------------------------------

                        if not payload[
                            "success"
                        ]:

                            diagnostic = payload[
                                "diagnostic"
                            ]

                            diagnostic[
                                "_outputDirectory"
                            ] = str(
                                payload[
                                    "outputDirectory"
                                ]
                            )

                            diagnostic_funds.append(
                                {
                                    "status":
                                        "DIAGNOSTIC_NO_CONFIRMED_RESULT",
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
                                    "factsheetUrl":
                                        payload.get(
                                            "factsheetUrl"
                                        ),
                                    "factsheetPageCount":
                                        diagnostic.get(
                                            "factsheetPageCount"
                                        ),
                                    "explicitSectionHeadings":
                                        len(
                                            diagnostic.get(
                                                "explicitSectionHeadings",
                                                [],
                                            )
                                        ),
                                    "rankCandidateCount":
                                        diagnostic.get(
                                            "rankCandidateCount",
                                            0,
                                        ),
                                    "percentageCandidateCount":
                                        diagnostic.get(
                                            "percentageCandidateCount",
                                            0,
                                        ),
                                    "keywordCandidateCount":
                                        diagnostic.get(
                                            "keywordCandidateCount",
                                            0,
                                        ),
                                    "spatialRankCandidateCount":
                                        diagnostic.get(
                                            "spatialRankCandidateCount",
                                            0,
                                        ),
                                    "spatialPercentageCandidateCount":
                                        diagnostic.get(
                                            "spatialPercentageCandidateCount",
                                            0,
                                        ),
                                    "recovery1Failure":
                                        excel_fund.get(
                                            "recovery1Failure"
                                        ),
                                    "diagnosticDirectory":
                                        str(
                                            payload[
                                                "outputDirectory"
                                            ]
                                        ),
                                    "generatedAtUtc":
                                        utc_now_iso(),
                                }
                            )

                            print(
                                "DIAGNOSTIC COMPLETE"
                            )

                            print(
                                "No single "
                                "structurally confirmed "
                                "holdings signature."
                            )

                            print(
                                "PDF forensic evidence saved to:"
                            )

                            print(
                                f"  {payload['outputDirectory']}"
                            )

                            break

                        result = payload[
                            "result"
                        ]

                        result[
                            "recovery1Failure"
                        ] = excel_fund.get(
                            "recovery1Failure"
                        )

                        # -----------------------------------------------------
                        # EXACT OFFICIAL PDF VERIFICATION
                        # -----------------------------------------------------

                        (
                            verified_pdf,
                            verified_text,
                            verified_section,
                            verified_holdings,
                        ) = (
                            verify_exact_against_official_pdf(
                                page,
                                result,
                            )
                        )

                        result[
                            "factsheetPageCount"
                        ] = len(
                            PdfReader(
                                BytesIO(
                                    verified_pdf
                                )
                            ).pages
                        )

                        save_success(
                            result,
                            verified_pdf,
                            verified_text,
                            verified_section,
                            payload[
                                "diagnostic"
                            ],
                        )

                        successful.append(
                            result
                        )

                        all_holdings.append(
                            result
                        )

                        print(
                            "RECOVERY 2 SUCCESS"
                        )

                        print(
                            f"Holdings: "
                            f"{len(verified_holdings)}"
                        )

                        print(
                            f"Method: "
                            f"{result.get('holdingsParser')}"
                        )

                        break

                    except Exception as error:

                        last_error = clean_text(
                            str(error)
                        )

                        print(
                            "Recovery 2 attempt failed: "
                            f"{last_error}"
                        )

                        if attempt < RETRY_COUNT:

                            time.sleep(
                                RETRY_DELAY_SECONDS
                            )

                else:

                    error_text = (
                        last_error
                        or
                        "Unknown Recovery 2 failure."
                    )

                    save_failure(
                        excel_fund,
                        error_text,
                    )

                    failed.append(
                        {
                            "status": "failed",
                            "excelRow": excel_row,
                            "prudentialUrl":
                                excel_fund[
                                    "prudentialUrl"
                                ],
                            "excelPruAccessName":
                                excel_fund.get(
                                    "pruAccessName"
                                ),
                            "error": error_text,
                            "recovery1Failure":
                                excel_fund.get(
                                    "recovery1Failure"
                                ),
                            "failedAtUtc":
                                utc_now_iso(),
                        }
                    )

                    print(
                        "RECOVERY 2 FAILED"
                    )

        finally:

            context.close()
            browser.close()

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    completed_at = utc_now_iso()

    total_holdings = sum(
        int(
            item.get(
                "topHoldingsCount",
                0,
            )
            or 0
        )
        for item in successful
    )

    if failed:

        status = "partial"

    elif diagnostic_funds:

        status = "diagnostic_complete"

    else:

        status = "success"

    run_summary = {
        "status": status,
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": str(
            EXCEL_FILE
        ),
        "recovery1RunSummary": str(
            RECOVERY_1_RUN_SUMMARY_FILE
        ),
        "recovery1FailureSource":
            recovery_1.get(
                "_recovery2FailureSource"
            ),
        "recovery1FailedFunds":
            recovery1_failed_count,
        "recovery1FailedRows":
            [
                item.get(
                    "excelRow"
                )
                for item in recovery_1.get(
                    "failedFundsDetail",
                    [],
                )
            ],
        "recovery2Universe":
            len(
                recovery_universe
            ),
        "recovery2UniverseRows":
            [
                item.get(
                    "excelRow"
                )
                for item in recovery_universe
            ],
        "successfulFunds":
            len(
                successful
            ),
        "diagnosticFunds":
            len(
                diagnostic_funds
            ),
        "failedFunds":
            len(
                failed
            ),
        "totalPublishedTopHoldings":
            total_holdings,
        "successfulFundsDetail":
            [
                {
                    "excelRow":
                        item.get(
                            "excelRow"
                        ),
                    "fundName":
                        item.get(
                            "fundName"
                        ),
                    "factsheetUrl":
                        item.get(
                            "factsheetUrl"
                        ),
                    "factsheetDocumentDate":
                        item.get(
                            "factsheetDocumentDate"
                        ),
                    "factsheetDataAsAt":
                        item.get(
                            "factsheetDataAsAt"
                        ),
                    "topHoldingsCount":
                        item.get(
                            "topHoldingsCount"
                        ),
                    "holdingsParser":
                        item.get(
                            "holdingsParser"
                        ),
                    "holdingsExtractionSource":
                        item.get(
                            "holdingsExtractionSource"
                        ),
                    "finalVerification":
                        item.get(
                            "finalVerification"
                        ),
                }
                for item in successful
            ],
        "diagnosticFundsDetail":
            diagnostic_funds,
        "failedFundsDetail":
            failed,
        "rules": {
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
            "spatialDiagnosticsEnabled": True,
            "adaptiveSectionDiscoveryEnabled": True,
            "exactFinalPdfVerification": True,
            "exactRankNameWeightSignatureRequired": True,
        },
    }

    save_json(
        RECOVERY_ALL_HOLDINGS_FILE,
        all_holdings,
    )

    save_json(
        RECOVERY_RUN_SUMMARY_FILE,
        run_summary,
    )

    # -------------------------------------------------------------------------
    # CONSOLE
    # -------------------------------------------------------------------------

    print()
    print()
    print(
        "=" * 78
    )
    print(
        "PRUDENTIAL TOP HOLDINGS RECOVERY 2 COMPLETE"
    )
    print(
        "=" * 78
    )

    print(
        f"Recovery 1 failure source: "
        f"{recovery_1.get('_recovery2FailureSource')}"
    )

    print(
        f"Recovery 1 failed funds: "
        f"{recovery1_failed_count}"
    )

    print(
        f"Recovery 2 universe: "
        f"{len(recovery_universe)}"
    )

    print(
        f"Successful with holdings: "
        f"{len(successful)}"
    )

    print(
        f"Diagnostic / unresolved: "
        f"{len(diagnostic_funds)}"
    )

    print(
        f"Failed technically: "
        f"{len(failed)}"
    )

    print(
        f"Total published holdings extracted: "
        f"{total_holdings}"
    )

    print()
    print(
        "Recovery 2 rows processed:"
    )

    print(
        "  "
        + ", ".join(
            str(
                item[
                    "excelRow"
                ]
            )
            for item in recovery_universe
        )
    )

    if diagnostic_funds:

        print()
        print(
            "DIAGNOSTIC FUNDS:"
        )

        for item in diagnostic_funds:

            print(
                f" - Row {item['excelRow']}: "
                f"ranks={item.get('rankCandidateCount', 0)}, "
                f"percentages={item.get('percentageCandidateCount', 0)}, "
                f"keywords={item.get('keywordCandidateCount', 0)}, "
                f"spatialRanks={item.get('spatialRankCandidateCount', 0)}, "
                f"spatialPercentages={item.get('spatialPercentageCandidateCount', 0)}"
            )

            print(
                f"   Diagnostics: "
                f"{item.get('diagnosticDirectory')}"
            )

    if failed:

        print()
        print(
            "RECOVERY 2 FAILED FUNDS:"
        )

        for failure in failed:

            print(
                f" - Row "
                f"{failure['excelRow']}: "
                f"{failure['error']}"
            )

    print()
    print(
        "Recovery 2 output:"
    )

    print(
        f" - {RECOVERY_ALL_HOLDINGS_FILE}"
    )

    print(
        f" - {RECOVERY_RUN_SUMMARY_FILE}"
    )

    print(
        f" - {RECOVERY_FUNDS_OUTPUT_DIR}"
    )

    print()
    print(
        "Done."
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
