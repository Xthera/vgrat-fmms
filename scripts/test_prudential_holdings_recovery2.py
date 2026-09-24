#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings RECOVERY 2 EXTRACTOR

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


HARD RULES
==========

- Recovery 1 failed funds control Recovery 2 universe.
- No baseline fallback.
- Excel Column A controls the master fund URL.
- No hardcoded fund count.
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
- Fixed-income lines may contain coupon/rate percentages.
- Fallback uses the LAST percentage on a logical line.
- Primary parser rejects ambiguous multiple-percentage lines.
- Spatial PDF recovery is permitted.
- No fuzzy holding matching.
- No arbitrary coordinate proximity pairing.
- No synthetic values.
- No estimated values.
- No interpolated values.
- Final verification downloads the official PDF again.
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
    )

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def save_json(path: Path, value) -> None:
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


def safe_filename(value: str) -> str:
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


def is_prudential_url(url: str) -> bool:
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


def ensure_prudential_url(url: str) -> str:
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


class HoldingsParseFailure(RuntimeError):
    pass


# =============================================================================
# RECOVERY 1 FAILURE EXTRACTION
# =============================================================================

def _normalize_failed_fund_entries(candidate) -> list[dict]:
    """
    Normalize possible Recovery 1 failure structures.

    Accepted forms include:

        [
            {"excelRow": 17, ...}
        ]

    or:

        {
            "17": {
                ...
            }
        }

    The function does not access baseline data.
    """

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

        if not isinstance(entry, dict):
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
                row = int(item[key])
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
    """
    Reconstruct Recovery 1 failures from Recovery 1's own output.

    ONLY:

        output_holdings_recovery/funds/*_failed/failure.json

    NEVER:

        output_holdings/
    """

    if not RECOVERY_FUNDS_OUTPUT_DIR.parent.exists():
        return []

    recovery1_funds_dir = Path(
        "output_holdings_recovery"
    ) / "funds"

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

        if not isinstance(payload, dict):
            continue

        if normalize_text(
            payload.get("status")
        ) != "failed":
            continue

        excel_row = payload.get("excelRow")

        if excel_row is None:

            match = re.match(
                r"^(\d+)_failed$",
                failure_file.parent.name,
            )

            if match:
                excel_row = int(match.group(1))

        if excel_row is None:
            raise RuntimeError(
                "Recovery 1 failure file has no identifiable "
                f"Excel row: {failure_file}"
            )

        try:
            excel_row = int(excel_row)
        except Exception as error:
            raise RuntimeError(
                "Invalid Excel row in Recovery 1 failure file: "
                f"{failure_file}"
            ) from error

        item = dict(payload)

        item["excelRow"] = excel_row

        item["_reconstructedFromRecovery1Output"] = True

        item["_recovery1FailureFile"] = str(
            failure_file
        )

        recovered.append(item)

    recovered.sort(
        key=lambda item: item["excelRow"]
    )

    return recovered


def load_recovery_1_run_summary() -> dict:
    """
    Load Recovery 1 state.

    Preferred order:

    1. failedFundsDetail
    2. alternative failure fields in Recovery 1 summary
    3. Recovery 1's own *_failed/failure.json files

    NEVER uses baseline output.
    """

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

    if not isinstance(payload, dict):

        raise RuntimeError(
            "Recovery 1 run summary is not a JSON object."
        )

    # -------------------------------------------------------------------------
    # 1. Preferred field
    # -------------------------------------------------------------------------

    normalized_failed = _normalize_failed_fund_entries(
        payload.get("failedFundsDetail")
    )

    source = ""

    if normalized_failed:

        source = (
            "recovery1_run_summary:"
            "failedFundsDetail"
        )

    # -------------------------------------------------------------------------
    # 2. Alternative fields
    # -------------------------------------------------------------------------

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

            candidate = payload.get(key)

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

    # -------------------------------------------------------------------------
    # 3. Recovery 1 own failure files
    # -------------------------------------------------------------------------

    if not normalized_failed:

        reconstructed = (
            load_recovery_1_failure_files()
        )

        if reconstructed:

            normalized_failed = reconstructed

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

    # -------------------------------------------------------------------------
    # Validate rows
    # -------------------------------------------------------------------------

    cleaned = []

    seen = set()

    for item in normalized_failed:

        row = item.get("excelRow")

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

    payload["failedFundsDetail"] = cleaned

    payload["_recovery2FailureSource"] = source

    return payload


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


# =============================================================================
# BUILD RECOVERY 2 UNIVERSE
# =============================================================================

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

    recovery_universe = []

    seen_rows = set()

    for failure in failed_entries:

        if not isinstance(failure, dict):

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

        seen_rows.add(excel_row)

        if excel_row not in excel_funds:

            raise RuntimeError(
                f"Recovery 1 failed row {excel_row} "
                "does not exist in Excel."
            )

        excel_fund = excel_funds[excel_row]

        excel_url = clean_text(
            excel_fund.get("prudentialUrl")
        )

        excel_name = clean_text(
            excel_fund.get("pruAccessName")
        )

        failure_url = clean_text(
            failure.get("prudentialUrl")
        )

        failure_name = clean_text(
            failure.get("excelPruAccessName")
            or failure.get("pruAccessName")
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

        recovery_universe.append(
            {
                "excelRow": excel_row,
                "prudentialUrl": excel_url,
                "pruAccessName": excel_name,
                "recovery1Failure": failure,
            }
        )

    return recovery_universe


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

    links = page.locator("a")

    candidates = []

    for index in range(
        links.count()
    ):

        anchor = links.nth(index)

        try:

            href = clean_text(
                anchor.get_attribute("href")
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

    if not body.startswith(b"%PDF"):

        raise RuntimeError(
            "Official factsheet response is not a PDF."
        )

    return body


# =============================================================================
# PDF TEXT
# =============================================================================

def extract_pdf_text(
    factsheet_bytes: bytes,
) -> tuple[str, int]:

    reader = PdfReader(
        BytesIO(factsheet_bytes)
    )

    pages = []

    for pdf_page in reader.pages:

        try:
            text = pdf_page.extract_text() or ""
        except Exception:
            text = ""

        pages.append(text)

    return (
        "\n".join(pages),
        len(reader.pages),
    )


def pdf_lines(text: str) -> list[str]:

    if not text:
        return []

    return [
        clean_text(line)
        for line in text.splitlines()
    ]


# =============================================================================
# FACTSHEET METADATA
# =============================================================================

def extract_data_as_at(text: str) -> str:

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
                match.group(1)
            )

    return ""


def extract_document_date(text: str) -> str:

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
                match.group(1)
            )

    return ""


# =============================================================================
# TOP HOLDINGS SECTION
# =============================================================================

def find_holdings_start(
    lines: list[str],
) -> int | None:

    for index, line in enumerate(lines):

        normalized = normalize_text(line)

        if re.match(
            r"^top\s+(?:10|ten)\s+holdings\b",
            normalized,
            flags=re.IGNORECASE,
        ):

            return index

    return None


def is_holdings_end(line: str) -> bool:

    normalized = normalize_text(line)

    if not normalized:
        return False

    markers = (
        "source",
        "inception date",
        "important information",
        "disclaimer",
        "past performance",
        "portfolio characteristics",
        "asset allocation",
        "page ",
    )

    return any(
        normalized.startswith(marker)
        for marker in markers
    )


def extract_holdings_section(
    full_text: str,
) -> tuple[str, str]:

    lines = pdf_lines(full_text)

    start = find_holdings_start(lines)

    if start is None:
        return "", "not_found"

    section = []

    for line in lines[start + 1:]:

        if is_holdings_end(line):
            break

        section.append(line)

    section_text = "\n".join(section).strip()

    if not section_text:
        return "", "published_empty"

    return section_text, "published"


# =============================================================================
# HOLDING NAME CLEANING
# =============================================================================

def clean_holding_name(value: str) -> str:

    name = clean_text(value)

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

    name = name.replace("|", " ")

    name = re.sub(
        r"\bNone\b",
        " ",
        name,
        flags=re.IGNORECASE,
    )

    name = clean_text(name).strip(" -|")

    if normalize_text(name) in {
        "",
        "none",
        "null",
        "-",
        "—",
    }:
        return ""

    return name


def clean_holding_fragment(value: str) -> str:

    text = clean_text(value)

    text = text.replace("|", " ")

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

    return clean_text(text)


def combine_holding_name_fragments(
    fragments: list[str],
) -> str:

    cleaned = []

    for fragment in fragments:

        fragment = clean_holding_fragment(
            fragment
        )

        if fragment:
            cleaned.append(fragment)

    return clean_holding_name(
        " ".join(cleaned)
    )


# =============================================================================
# PERCENTAGES
# =============================================================================

PERCENTAGE_RE = re.compile(
    r"(?<![\d.])"
    r"(\d+(?:\.\d+)?)"
    r"\s*%"
)


def percentage_matches(line: str):
    return list(
        PERCENTAGE_RE.finditer(line)
    )


def find_last_percentage_in_line(line: str):

    matches = percentage_matches(line)

    if not matches:
        return None

    match = matches[-1]

    return (
        float(match.group(1)),
        match.group(0),
        match.start(),
        match.end(),
    )


# =============================================================================
# RANKS
# =============================================================================

def extract_leading_rank(line: str):

    text = clean_text(line)

    match = re.match(
        r"^(\d{1,2})[.)]?\s+(.*)$",
        text,
    )

    if match:

        rank = int(match.group(1))

        if 1 <= rank <= 99:

            return (
                rank,
                clean_text(
                    match.group(2)
                ),
            )

    return None, text


# =============================================================================
# NOISE
# =============================================================================

def is_holding_header_or_noise(
    line: str,
) -> bool:

    normalized = normalize_text(line)

    if not normalized:
        return True

    if normalized in {
        "top 10 holdings",
        "top ten holdings",
        "holding",
        "holdings",
        "weight",
        "percentage",
        "%",
        "name",
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
# PDF LINE RECONSTRUCTION
# =============================================================================

def merge_hyphenated_line_breaks(
    lines: list[str],
) -> list[str]:

    merged = []

    index = 0

    while index < len(lines):

        current = clean_text(lines[index])

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
            merged.append(current)

        index += 1

    return merged


def build_logical_holding_lines(
    section_text: str,
) -> list[str]:

    lines = pdf_lines(section_text)

    lines = merge_hyphenated_line_breaks(lines)

    return [
        line
        for line in lines
        if clean_text(line)
    ]


# =============================================================================
# VALIDATION
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

    expected = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    actual = [
        item.get("rank")
        for item in holdings
    ]

    if actual != expected:

        raise HoldingsParseFailure(
            "Holding ranks are not sequential. "
            f"Parsed={actual}; Expected={expected}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding.get("name")
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


# =============================================================================
# PRIMARY PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:

    if not clean_text(section_text):

        raise HoldingsParseFailure(
            "Top 10 holdings section is empty."
        )

    logical_lines = build_logical_holding_lines(
        section_text
    )

    holdings = []

    pending_fragments = []

    pending_rank = None

    for raw_line in logical_lines:

        line = clean_text(raw_line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, remainder = extract_leading_rank(line)

        if rank is not None:

            if pending_fragments:

                raise HoldingsParseFailure(
                    "New holding rank appeared before "
                    "the previous holding received a percentage."
                )

            pending_rank = rank
            line = remainder

        matches = percentage_matches(line)

        if not matches:

            if pending_rank is not None or holdings:

                fragment = clean_holding_fragment(line)

                if fragment:
                    pending_fragments.append(fragment)

            continue

        if len(matches) != 1:

            raise HoldingsParseFailure(
                "Primary parser encountered multiple "
                "percentages on one logical line."
            )

        match = matches[0]

        percentage = float(match.group(1))

        percentage_text = match.group(0)

        fragment = clean_holding_fragment(
            line[:match.start()]
        )

        if fragment:
            pending_fragments.append(fragment)

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

    return validate_holdings(holdings)


# =============================================================================
# FALLBACK PARSER
# =============================================================================

def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:

    if not clean_text(section_text):

        raise HoldingsParseFailure(
            "Top 10 holdings section is empty."
        )

    logical_lines = build_logical_holding_lines(
        section_text
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
                f"Fallback parser produced more than "
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

        line = clean_text(raw_line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, remainder = extract_leading_rank(line)

        if rank is not None:

            if pending_fragments:

                raise HoldingsParseFailure(
                    "Fallback parser encountered a new "
                    "rank before the previous holding "
                    "received a published percentage."
                )

            pending_rank = rank
            line = remainder

        last_percentage = find_last_percentage_in_line(
            line
        )

        if last_percentage is None:

            fragment = clean_holding_fragment(line)

            if fragment:
                pending_fragments.append(fragment)

            continue

        percentage, percentage_text, start, end = (
            last_percentage
        )

        fragment = clean_holding_fragment(
            line[:start]
        )

        if fragment:
            pending_fragments.append(fragment)

        commit(
            percentage,
            percentage_text,
        )

    if pending_fragments:

        raise HoldingsParseFailure(
            "Fallback parser found trailing holding text "
            "without a published percentage."
        )

    return validate_holdings(holdings)


# =============================================================================
# SIGNATURE
# =============================================================================

def holding_signature(
    holdings: list[dict],
) -> list[tuple]:

    return [
        (
            item.get("rank"),
            normalize_text(item.get("name")),
            float(item.get("weightPercent")),
        )
        for item in holdings
    ]


# =============================================================================
# SPATIAL PDF EXTRACTION
# =============================================================================

def positioned_pdf_lines(
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

                value = clean_text(text)

                if not value:
                    return

                try:

                    x = float(tm[4])
                    y = float(tm[5])

                except Exception:
                    return

                words.append(
                    {
                        "page": page_number,
                        "text": value,
                        "x": x,
                        "y": y,
                        "fontSize": float(
                            font_size or 0
                        ),
                    }
                )

            pdf_page.extract_text(
                visitor_text=visitor_text
            )

        except Exception:

            words = []

        positioned.extend(words)

    return positioned


def find_spatial_heading(
    positioned: list[dict],
):

    for item in positioned:

        normalized = normalize_text(
            item.get("text")
        )

        if re.match(
            r"^top\s+(?:10|ten)\s+holdings\b",
            normalized,
            flags=re.IGNORECASE,
        ):

            return item

    return None


def build_spatial_candidates(
    positioned: list[dict],
    heading: dict,
) -> list[dict]:

    page = heading.get("page")
    heading_y = heading.get("y")

    if page is None:
        return []

    page_items = [
        item
        for item in positioned
        if item.get("page") == page
    ]

    if heading_y is not None:

        page_items = [
            item
            for item in page_items
            if item.get("y", 0) < heading_y
        ]

    rows = {}

    for item in page_items:

        y = round(
            float(item.get("y", 0)),
            1,
        )

        rows.setdefault(y, []).append(item)

    candidates = []

    for y in sorted(
        rows,
        reverse=True,
    ):

        row = sorted(
            rows[y],
            key=lambda item: item.get("x", 0),
        )

        text = clean_text(
            " ".join(
                item.get("text", "")
                for item in row
            )
        )

        if not text:
            continue

        if is_holding_header_or_noise(text):
            continue

        if percentage_matches(text):

            candidates.append(
                {
                    "page": page,
                    "y": y,
                    "text": text,
                    "items": row,
                }
            )

    return candidates


def choose_spatial_left_column(
    candidates: list[dict],
) -> list[dict]:

    if not candidates:
        return []

    weight_items = []

    for candidate in candidates:

        for item in candidate.get(
            "items",
            [],
        ):

            if percentage_matches(
                item.get("text", "")
            ):

                weight_items.append(item)

    if not weight_items:
        return []

    boundary = max(
        item.get("x", 0)
        for item in weight_items
    )

    selected = []

    for candidate in candidates:

        left_items = [
            item
            for item in candidate.get(
                "items",
                [],
            )
            if item.get("x", 0) < boundary
        ]

        if not left_items:
            continue

        name_text = clean_text(
            " ".join(
                item.get("text", "")
                for item in left_items
            )
        )

        if not name_text:
            continue

        selected.append(
            {
                "y": candidate.get("y"),
                "name": name_text,
                "fullText": candidate.get("text"),
            }
        )

    return selected


def spatial_lines_to_section(
    spatial_lines: list[dict],
) -> str:

    return "\n".join(
        clean_text(
            item.get(
                "fullText",
                "",
            )
        )
        for item in spatial_lines
        if clean_text(
            item.get(
                "fullText",
                "",
            )
        )
    )


def extract_spatial_fallback(
    factsheet_bytes: bytes,
) -> tuple[str, list[dict]]:

    positioned = positioned_pdf_lines(
        factsheet_bytes
    )

    if not positioned:

        raise HoldingsParseFailure(
            "No positioned PDF text was available."
        )

    heading = find_spatial_heading(
        positioned
    )

    if heading is None:

        raise HoldingsParseFailure(
            "Spatial PDF fallback could not locate "
            "the Top 10 Holdings heading."
        )

    candidates = build_spatial_candidates(
        positioned,
        heading,
    )

    selected = choose_spatial_left_column(
        candidates
    )

    if not selected:

        raise HoldingsParseFailure(
            "Spatial PDF fallback could not reconstruct "
            "holding lines."
        )

    section_text = spatial_lines_to_section(
        selected
    )

    if not section_text:

        raise HoldingsParseFailure(
            "Spatial PDF fallback produced an empty section."
        )

    holdings = parse_holdings_fallback(
        section_text
    )

    return section_text, holdings


# =============================================================================
# RECOVERY EXTRACTION ENGINE
# =============================================================================

def extract_with_recovery_engine(
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
) -> tuple[list[dict], str, str, str]:

    primary_error = ""
    fallback_error = ""

    try:

        holdings = parse_holdings(
            section_text
        )

        return (
            holdings,
            "primary",
            primary_error,
            fallback_error,
        )

    except Exception as error:

        primary_error = clean_text(
            str(error)
        )

    try:

        holdings = parse_holdings_fallback(
            section_text
        )

        return (
            holdings,
            "fallback",
            primary_error,
            fallback_error,
        )

    except Exception as error:

        fallback_error = clean_text(
            str(error)
        )

    try:

        _spatial_section, holdings = (
            extract_spatial_fallback(
                factsheet_bytes
            )
        )

        return (
            holdings,
            "spatial_fallback",
            primary_error,
            fallback_error,
        )

    except Exception as error:

        spatial_error = clean_text(
            str(error)
        )

        raise HoldingsParseFailure(
            "All Recovery 2 extraction strategies failed. "
            f"Primary: {primary_error}; "
            f"Fallback: {fallback_error}; "
            f"Spatial: {spatial_error}"
        )


# =============================================================================
# FUND PAGE NAME
# =============================================================================

def extract_fund_page_name(page) -> str:

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
# RECOVER SINGLE FUND
# =============================================================================

def recover_single_fund(
    page,
    excel_fund: dict,
) -> dict:

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

    full_text, pdf_page_count = (
        extract_pdf_text(
            factsheet_bytes
        )
    )

    section_text, section_status = (
        extract_holdings_section(
            full_text
        )
    )

    if section_status != "published":

        if section_status == "not_found":

            raise HoldingsParseFailure(
                "Official factsheet has no confirmed "
                "Top 10 Holdings section."
            )

        raise HoldingsParseFailure(
            "Official factsheet Top 10 Holdings section "
            "was empty."
        )

    (
        holdings,
        parser_used,
        primary_error,
        fallback_error,
    ) = extract_with_recovery_engine(
        factsheet_bytes,
        full_text,
        section_text,
    )

    holdings = validate_holdings(
        holdings
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
        "holdingsSectionStatus": "published",
        "topHoldingsCount": len(holdings),
        "topHoldings": holdings,
        "holdingsParser": parser_used,
        "primaryParserError": primary_error,
        "fallbackParserError": fallback_error,
        "holdingsSectionExtraction": (
            "pdf_coordinates"
            if parser_used == "spatial_fallback"
            else "pdf_text"
        ),
        "rules": {
            "recovery1FailedFundsOnly": True,
            "officialPrudentialSourceOnly": True,
            "officialFactsheetOnly": True,
            "maximumHoldings": MAX_HOLDINGS,
            "publishedCountUsedExactly": True,
            "fewerThanTenAllowed": True,
            "noForcedTenEntries": True,
            "noInferredHoldings": True,
            "noFabricatedPercentages": True,
            "duplicateHoldingNamesAllowed": True,
            "duplicateHoldingPercentagesAllowed": True,
            "multilineHoldingNamesSupported": True,
            "wrappedPdfLinesJoined": True,
            "fallbackUsesLastPercentageOnLogicalLine": True,
            "spatialFallbackEnabled": True,
            "noFuzzyHoldingMatching": True,
            "noCoordinateProximityPairing": True,
            "noSyntheticValues": True,
            "exactFinalPdfVerification": True,
            "exactRankNameWeightSignatureRequired": True,
        },
    }

    return {
        "result": result,
        "factsheetBytes": factsheet_bytes,
        "fullText": full_text,
        "sectionText": section_text,
    }


# =============================================================================
# FINAL EXACT VERIFICATION
# =============================================================================

def verify_exact_against_official_pdf(
    page,
    result: dict,
) -> tuple[bytes, str, str, list[dict]]:

    factsheet_url = ensure_prudential_url(
        result["factsheetUrl"]
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

    full_text, _page_count = extract_pdf_text(
        factsheet_bytes
    )

    section_text, section_status = (
        extract_holdings_section(
            full_text
        )
    )

    if section_status != "published":

        raise RuntimeError(
            "Factsheet holdings section disappeared "
            "during final verification."
        )

    parser_used = result.get(
        "holdingsParser"
    ) or "primary"

    if parser_used == "primary":

        verified_holdings = parse_holdings(
            section_text
        )

    elif parser_used == "fallback":

        verified_holdings = parse_holdings_fallback(
            section_text
        )

    elif parser_used == "spatial_fallback":

        spatial_section, verified_holdings = (
            extract_spatial_fallback(
                factsheet_bytes
            )
        )

        section_text = spatial_section

    else:

        raise RuntimeError(
            "Unknown parser used during extraction: "
            f"{parser_used}"
        )

    original_holdings = (
        result.get("topHoldings") or []
    )

    if len(verified_holdings) != len(
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

    result["topHoldings"] = verified_holdings

    result["topHoldingsCount"] = len(
        verified_holdings
    )

    result["finalVerification"] = {
        "status": "verified",
        "verifiedAtUtc": utc_now_iso(),
        "exactSignatureMatch": True,
        "verifiedHoldingCount": len(
            verified_holdings
        ),
    }

    return (
        factsheet_bytes,
        full_text,
        section_text,
        verified_holdings,
    )


# =============================================================================
# OUTPUT HELPERS
# =============================================================================

def recovery_directory(
    result: dict,
) -> Path:

    excel_row = int(
        result["excelRow"]
    )

    identifier = safe_filename(
        result.get("fundName")
        or result.get("pageTitle")
        or urlparse(
            result.get("finalUrl")
            or result.get("prudentialUrl")
        ).path.rstrip("/").split("/")[-1]
        or f"fund_{excel_row}"
    )

    return (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )


def save_success(
    result: dict,
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
) -> Path:

    directory = recovery_directory(
        result
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
        {
            "recoveryStage": "Recovery 2",
            "recovery1Failure": result.get(
                "recovery1Failure"
            ),
            "holdingsParser": result.get(
                "holdingsParser"
            ),
            "primaryParserError": result.get(
                "primaryParserError"
            ),
            "fallbackParserError": result.get(
                "fallbackParserError"
            ),
            "finalVerification": result.get(
                "finalVerification"
            ),
            "savedAtUtc": utc_now_iso(),
        },
    )

    save_json(
        directory / "metadata.json",
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
            "factsheetDocumentDate": result.get(
                "factsheetDocumentDate"
            ),
            "factsheetDataAsAt": result.get(
                "factsheetDataAsAt"
            ),
            "topHoldingsCount": result.get(
                "topHoldingsCount"
            ),
            "status": result.get(
                "status"
            ),
            "recoveryStage": "Recovery 2",
            "savedAtUtc": utc_now_iso(),
        },
    )

    return directory


def save_no_holdings_section(
    excel_fund: dict,
    error_text: str,
) -> Path:

    excel_row = int(
        excel_fund["excelRow"]
    )

    directory = (
        RECOVERY_FUNDS_OUTPUT_DIR
        / f"{excel_row}_no_holdings_section"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory / "failure.json",
        {
            "status": "NO_HOLDINGS_SECTION",
            "excelRow": excel_row,
            "prudentialUrl": excel_fund[
                "prudentialUrl"
            ],
            "excelPruAccessName": excel_fund.get(
                "pruAccessName"
            ),
            "error": clean_text(error_text),
            "recoveryStage": "Recovery 2",
            "failedAtUtc": utc_now_iso(),
        },
    )

    return directory


def save_failure(
    excel_fund: dict,
    error_text: str,
    recovery1_failure=None,
) -> Path:

    excel_row = int(
        excel_fund["excelRow"]
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
        directory / "failure.json",
        {
            "status": "failed",
            "excelRow": excel_row,
            "prudentialUrl": excel_fund[
                "prudentialUrl"
            ],
            "excelPruAccessName": excel_fund.get(
                "pruAccessName"
            ),
            "error": clean_text(error_text),
            "recoveryStage": "Recovery 2",
            "recovery1Failure": recovery1_failure,
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

    if recovery_universe:

        print(
            "Recovery 1 failed rows:"
        )

        print(
            "  "
            + ", ".join(
                str(
                    item["excelRow"]
                )
                for item in recovery_universe
            )
        )

    print(
        f"Recovery 2 universe: "
        f"{len(recovery_universe)}"
    )

    # -------------------------------------------------------------------------
    # NOTHING TO RECOVER
    # -------------------------------------------------------------------------

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
                "excelFile": str(EXCEL_FILE),
                "recovery1RunSummary": str(
                    RECOVERY_1_RUN_SUMMARY_FILE
                ),
                "recovery1FailureSource": recovery_1.get(
                    "_recovery2FailureSource"
                ),
                "recovery1ExcelFundUniverse": recovery_1.get(
                    "excelFundUniverse"
                ),
                "recovery1FailedFunds": recovery1_failed_count,
                "recovery2Universe": 0,
                "successfulFunds": 0,
                "failedFunds": 0,
                "successfulFundsDetail": [],
                "failedFundsDetail": [],
                "rules": {
                    "recovery1FailedFundsOnly": True,
                    "noBaselineFallback": True,
                    "officialPrudentialSingaporeOnly": True,
                    "officialFactsheetOnly": True,
                    "noThirdPartyHoldings": True,
                    "noInferredHoldings": True,
                    "noFabricatedHoldings": True,
                    "noSyntheticValues": True,
                    "exactFinalPdfVerification": True,
                },
            },
        )

        print()
        print(
            "Recovery 1 has no failed funds."
        )
        print(
            "Recovery 2 has nothing to recover."
        )

        return 0

    # -------------------------------------------------------------------------
    # PROCESS FUNDS
    # -------------------------------------------------------------------------

    successful = []

    failed = []

    no_holdings_section = []

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
                    excel_fund["excelRow"]
                )

                print()
                print(
                    "=" * 72
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
                    "=" * 72
                )

                result_payload = None
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

                        result_payload = (
                            recover_single_fund(
                                page,
                                excel_fund,
                            )
                        )

                        result = (
                            result_payload[
                                "result"
                            ]
                        )

                        result[
                            "recovery1Failure"
                        ] = excel_fund.get(
                            "recovery1Failure"
                        )

                        (
                            verified_pdf,
                            verified_text,
                            verified_section,
                            verified_holdings,
                        ) = verify_exact_against_official_pdf(
                            page,
                            result,
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
                        )

                        successful.append(result)
                        all_holdings.append(result)

                        print(
                            "RECOVERY 2 SUCCESS"
                        )

                        print(
                            f"Holdings: "
                            f"{len(verified_holdings)}"
                        )

                        print(
                            f"Parser: "
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

                    if (
                        "no confirmed Top 10 Holdings section"
                        in error_text
                    ):

                        save_no_holdings_section(
                            excel_fund,
                            error_text,
                        )

                        no_holdings_section.append(
                            {
                                "excelRow": excel_row,
                                "prudentialUrl": (
                                    excel_fund[
                                        "prudentialUrl"
                                    ]
                                ),
                                "excelPruAccessName": (
                                    excel_fund.get(
                                        "pruAccessName"
                                    )
                                ),
                                "error": error_text,
                                "recovery1Failure": (
                                    excel_fund.get(
                                        "recovery1Failure"
                                    )
                                ),
                            }
                        )

                    else:

                        save_failure(
                            excel_fund,
                            error_text,
                            excel_fund.get(
                                "recovery1Failure"
                            ),
                        )

                        failed.append(
                            {
                                "status": "failed",
                                "excelRow": excel_row,
                                "prudentialUrl": (
                                    excel_fund[
                                        "prudentialUrl"
                                    ]
                                ),
                                "excelPruAccessName": (
                                    excel_fund.get(
                                        "pruAccessName"
                                    )
                                ),
                                "error": error_text,
                                "recovery1Failure": (
                                    excel_fund.get(
                                        "recovery1Failure"
                                    )
                                ),
                                "failedAtUtc": utc_now_iso(),
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

    run_summary = {
        "status": (
            "success"
            if not failed
            else "partial"
        ),

        "startedAtUtc":
            started_at,

        "completedAtUtc":
            completed_at,

        "excelFile":
            str(EXCEL_FILE),

        "recovery1RunSummary":
            str(
                RECOVERY_1_RUN_SUMMARY_FILE
            ),

        "recovery1FailureSource":
            recovery_1.get(
                "_recovery2FailureSource"
            ),

        "recovery1ExcelFundUniverse":
            recovery_1.get(
                "excelFundUniverse"
            ),

        "recovery1FailedFunds":
            recovery1_failed_count,

        "recovery1FailedRows":
            [
                item.get("excelRow")
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
                item.get("excelRow")
                for item in recovery_universe
            ],

        "successfulFunds":
            len(successful),

        "noHoldingsSectionFunds":
            len(no_holdings_section),

        "failedFunds":
            len(failed),

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
                }
                for item in successful
            ],

        "noHoldingsSectionDetail":
            no_holdings_section,

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
            "maximumHoldings": MAX_HOLDINGS,
            "publishedCountUsedExactly": True,
            "fewerThanTenHoldingsAllowed": True,
            "noForcedTenEntries": True,
            "duplicateHoldingNamesAllowed": True,
            "duplicateHoldingPercentagesAllowed": True,
            "multilineHoldingNamesSupported": True,
            "hyphenatedLineWrapReconstruction": True,
            "fallbackUsesLastPercentageOnLogicalLine": True,
            "spatialFallbackEnabled": True,
            "noFuzzyHoldingMatching": True,
            "noCoordinateProximityPairing": True,
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
    # FINAL CONSOLE
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
        f"No Top Holdings section: "
        f"{len(no_holdings_section)}"
    )

    print(
        f"Failed: "
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
                item["excelRow"]
            )
            for item in recovery_universe
        )
    )

    print()
    print(
        "Recovery source:"
    )

    print(
        f" - {RECOVERY_1_RUN_SUMMARY_FILE}"
    )

    if recovery_1.get(
        "_recovery2FailureSource"
    ) == "recovery1_output_failure_files":

        print(
            " - Recovery 1 *_failed/failure.json files"
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

        raise SystemExit(130)
