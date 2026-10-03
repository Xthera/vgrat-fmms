#!/usr/bin/env python3

"""
VGrat FMS - BUILD FUNDS DATA

Builds:

    data/funds.json
    data/bid_history.json

from the existing Prudential extraction pipelines.

============================================================
MASTER FUND UNIVERSE
============================================================

Funds Links.xlsm is the controlling universe.

    Column A = Prudential URL
    Column B = exact PruAccess Fund Name

Research Funds.xlsx is NOT a fund universe.

It is an extraction-only research workbook.

    Worksheet 1:
        Fund Research
        Per-fund Geography / Sector research

    Worksheet 2:
        Geography master categories
        AUTHORITATIVE UI Geography filter values

    Worksheet 3:
        Sector master categories
        AUTHORITATIVE UI Sector filter values

The master categories are copied into funds.json and are
NOT dynamically generated from the funds.

============================================================
IMPORTANT
============================================================

This script does NOT perform AI research.

Research Funds.xlsx is assumed to already contain the
researched values.

Existing holdings and BID pipelines remain the source for:

    Top Holdings
    Historical BID observations

Dividend fields are taken from the Prudential fund extraction.

The source field is:

    dividendRate
    dividendUnit

hasDividend is derived from dividendRate.

Therefore:

    dividendRate exists -> hasDividend = true

The upstream hasDividend flag is NOT trusted.

============================================================
OUTPUT
============================================================

data/funds.json

    {
      "schemaVersion": 1,
      "generatedAt": "...",
      "research": {
        "source": "Research Funds.xlsx",
        "fundResearchWorksheet": "...",
        "masterCategoryWorksheets": {
          "geography": "...",
          "sector": "..."
        },
        "masterCategories": {
          "geography": [...],
          "sector": [...]
        }
      },
      "summary": {...},
      "funds": [...]
    }

data/bid_history.json

    {
      "schemaVersion": 1,
      "generatedAt": "...",
      "funds": [...]
    }

============================================================
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

MASTER_EXCEL_FILE = ROOT / "Funds Links.xlsm"
RESEARCH_EXCEL_FILE = ROOT / "Research Funds.xlsx"

RESEARCH_FUND_SHEET_NAME = "Fund Research"

# Worksheet positions are zero-based internally.
#
# Worksheet 1 = Fund Research
# Worksheet 2 = Geography master categories
# Worksheet 3 = Sector master categories
#
# Therefore:
RESEARCH_GEOGRAPHY_SHEET_INDEX = 1
RESEARCH_SECTOR_SHEET_INDEX = 2

DATA_DIR = ROOT / "data"

FUNDS_OUTPUT_FILE = DATA_DIR / "funds.json"
BID_HISTORY_OUTPUT_FILE = DATA_DIR / "bid_history.json"

# Existing holdings pipeline outputs.
BASELINE_HOLDINGS_FILE = ROOT / "output_holdings" / "all_holdings.json"

RECOVERY_DIRECTORIES = [
    ROOT / "output_holdings_recovery",
    ROOT / "output_holdings_recovery_2",
    ROOT / "output_holdings_recovery_3",
]

# Existing PruAccess production directory.
PRUACCESS_PRODUCTION_DIR = ROOT / "output_pruaccess" / "production"

# Existing fund-level PruAccess directory.
PRUACCESS_FUNDS_DIR = ROOT / "output_pruaccess" / "funds"

SCHEMA_VERSION = 1

MAX_HOLDINGS = 10

ALLOW_UNRESOLVED = (
    os.environ.get("ALLOW_UNRESOLVED", "false").strip().lower()
    in {"1", "true", "yes", "y"}
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(value: Any) -> str:
    """
    Convert a value to clean text.

    None -> ""
    Strings are stripped and repeated whitespace collapsed.
    """
    if value is None:
        return ""

    if isinstance(value, str):
        return re.sub(r"\s+", " ", value).strip()

    return str(value).strip()


def clean_raw_text(value: Any) -> str:
    """
    Convert a value to string while preserving internal content
    as much as practical.

    Useful for fields such as dividendRate.
    """
    if value is None:
        return ""

    if isinstance(value, str):
        return value.strip()

    return str(value).strip()


def utc_now_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    temporary = path.with_suffix(path.suffix + ".tmp")

    with temporary.open("w", encoding="utf-8") as handle:
        json.dump(
            payload,
            handle,
            ensure_ascii=False,
            indent=2,
        )
        handle.write("\n")

    temporary.replace(path)


def normalize_match_key(value: Any) -> str:
    """
    Normalization used for matching fund names.

    This is deliberately conservative.

    It does NOT remove meaningful punctuation or words.
    """
    text = clean_text(value)

    text = text.replace("\u00a0", " ")

    return re.sub(r"\s+", " ", text).strip().casefold()


def normalize_category_key(value: Any) -> str:
    """
    Normalization used only for category de-duplication.
    """
    return normalize_match_key(value)


def is_empty_value(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, str):
        return not value.strip()

    return False


# ============================================================
# EXCEL HELPERS
# ============================================================

def open_excel_workbook(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")

    keep_vba = path.suffix.lower() == ".xlsm"

    return load_workbook(
        filename=path,
        data_only=True,
        read_only=False,
        keep_vba=keep_vba,
    )


def find_header_row(
    worksheet,
    required_headers: Iterable[str],
    max_scan_rows: int = 10,
) -> Tuple[int, Dict[str, int]]:
    """
    Find a header row by looking for the supplied header names.

    Returns:

        (row_number, normalized_header -> column_number)

    Raises if the headers cannot be found.
    """

    required = {
        normalize_match_key(header)
        for header in required_headers
    }

    max_row = min(worksheet.max_row, max_scan_rows)

    for row_number in range(1, max_row + 1):
        mapping: Dict[str, int] = {}

        for cell in worksheet[row_number]:
            value = normalize_match_key(cell.value)

            if value:
                mapping[value] = cell.column

        if required.issubset(mapping.keys()):
            return row_number, mapping

    raise ValueError(
        f"Could not locate required headers in worksheet "
        f"'{worksheet.title}': {sorted(required)}"
    )


def find_optional_column(
    worksheet,
    candidate_headers: Iterable[str],
    max_scan_rows: int = 10,
) -> Optional[Tuple[int, int]]:
    """
    Locate an optional column.

    Returns:

        (header_row, column_number)

    or None.
    """

    candidates = {
        normalize_match_key(value)
        for value in candidate_headers
    }

    max_row = min(worksheet.max_row, max_scan_rows)

    for row_number in range(1, max_row + 1):
        for cell in worksheet[row_number]:
            value = normalize_match_key(cell.value)

            if value in candidates:
                return row_number, cell.column

    return None


# ============================================================
# MASTER FUND UNIVERSE
# ============================================================

def read_master_funds() -> List[Dict[str, Any]]:
    """
    Read Funds Links.xlsm.

    Column A:
        Prudential URL

    Column B:
        Exact PruAccess Fund Name

    The workbook controls the fund universe.
    """

    workbook = open_excel_workbook(MASTER_EXCEL_FILE)

    try:
        worksheet = workbook.active

        funds: List[Dict[str, Any]] = []

        for row_number in range(2, worksheet.max_row + 1):
            prudential_url = clean_text(
                worksheet.cell(row=row_number, column=1).value
            )

            pruaccess_name = clean_text(
                worksheet.cell(row=row_number, column=2).value
            )

            if not prudential_url and not pruaccess_name:
                continue

            if not prudential_url:
                continue

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": prudential_url,
                    "pruAccessName": pruaccess_name,
                }
            )

        return funds

    finally:
        workbook.close()


# ============================================================
# RESEARCH WORKBOOK
# ============================================================

def find_research_headers(worksheet) -> Tuple[int, Dict[str, int]]:
    """
    Locate the Fund Research headers.

    Required:

        Fund Name

    Optional:

        Geographic 1
        Geographic 2
        Sector 1
        Sector 2

    Several common header spellings are accepted.
    """

    fund_name_candidates = {
        "fund name",
        "fund",
        "name",
        "pruaccess fund name",
    }

    geographic_1_candidates = {
        "geographic 1",
        "geography 1",
        "geographical 1",
        "geographic1",
        "geography1",
    }

    geographic_2_candidates = {
        "geographic 2",
        "geography 2",
        "geographical 2",
        "geographic2",
        "geography2",
    }

    sector_1_candidates = {
        "sector 1",
        "sector1",
    }

    sector_2_candidates = {
        "sector 2",
        "sector2",
    }

    max_scan_rows = min(worksheet.max_row, 10)

    for row_number in range(1, max_scan_rows + 1):

        normalized_row: Dict[str, int] = {}

        for cell in worksheet[row_number]:
            value = normalize_match_key(cell.value)

            if value:
                normalized_row[value] = cell.column

        fund_column = None

        for candidate in fund_name_candidates:
            if candidate in normalized_row:
                fund_column = normalized_row[candidate]
                break

        if fund_column is None:
            continue

        result: Dict[str, int] = {
            "fundName": fund_column,
        }

        for key, candidates in [
            ("geographic1", geographic_1_candidates),
            ("geographic2", geographic_2_candidates),
            ("sector1", sector_1_candidates),
            ("sector2", sector_2_candidates),
        ]:
            for candidate in candidates:
                if candidate in normalized_row:
                    result[key] = normalized_row[candidate]
                    break

        return row_number, result

    raise ValueError(
        f"Could not locate Fund Name column in worksheet "
        f"'{worksheet.title}'."
    )


def read_research_funds(
    workbook,
) -> Dict[str, Dict[str, Any]]:
    """
    Read worksheet 1 / Fund Research.

    Returns a dictionary keyed by normalized fund name.
    """

    if RESEARCH_FUND_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(
            f"Required research worksheet "
            f"'{RESEARCH_FUND_SHEET_NAME}' was not found. "
            f"Available worksheets: {workbook.sheetnames}"
        )

    worksheet = workbook[RESEARCH_FUND_SHEET_NAME]

    header_row, columns = find_research_headers(worksheet)

    results: Dict[str, Dict[str, Any]] = {}

    for row_number in range(
        header_row + 1,
        worksheet.max_row + 1,
    ):
        fund_name = clean_text(
            worksheet.cell(
                row=row_number,
                column=columns["fundName"],
            ).value
        )

        if not fund_name:
            continue

        record = {
            "fundName": fund_name,
            "geographic1": "",
            "geographic2": "",
            "sector1": "",
            "sector2": "",
        }

        for field in (
            "geographic1",
            "geographic2",
            "sector1",
            "sector2",
        ):
            column = columns.get(field)

            if column is not None:
                record[field] = clean_text(
                    worksheet.cell(
                        row=row_number,
                        column=column,
                    ).value
                )

        key = normalize_match_key(fund_name)

        # Last occurrence wins if a duplicate exists.
        results[key] = record

    return results


# ============================================================
# MASTER CATEGORY EXTRACTION
# ============================================================

GENERIC_CATEGORY_HEADERS = {
    "category",
    "categories",
    "master category",
    "master categories",
    "filter",
    "filters",
    "filter category",
    "filter categories",
    "name",
    "names",
    "value",
    "values",
}


def locate_category_column(
    worksheet,
    category_type: str,
) -> Tuple[int, int]:
    """
    Locate the category column in a master-category worksheet.

    We first look for a semantically appropriate header.

    If no specific header is found, we fall back to the first
    non-empty column in the first few rows.

    This allows the workbook to use slightly different labels
    without making the category extraction fragile.
    """

    if category_type == "geography":
        candidates = {
            "geography",
            "geographies",
            "geographic",
            "geographical",
            "geographic category",
            "geography category",
            "geography categories",
            "master geography",
            "master geographies",
        }
    else:
        candidates = {
            "sector",
            "sectors",
            "sector category",
            "sector categories",
            "master sector",
            "master sectors",
        }

    candidates = {
        normalize_match_key(value)
        for value in candidates
    }

    max_scan_rows = min(worksheet.max_row, 10)

    # First try explicit semantic headers.
    for row_number in range(1, max_scan_rows + 1):
        for cell in worksheet[row_number]:
            normalized = normalize_match_key(cell.value)

            if normalized in candidates:
                return row_number, cell.column

    # Then generic category headers.
    for row_number in range(1, max_scan_rows + 1):
        for cell in worksheet[row_number]:
            normalized = normalize_match_key(cell.value)

            if normalized in GENERIC_CATEGORY_HEADERS:
                return row_number, cell.column

    # Final fallback:
    # find the first column with actual content.
    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        for row_number in range(
            1,
            min(worksheet.max_row, 10) + 1,
        ):
            value = clean_text(
                worksheet.cell(
                    row=row_number,
                    column=column_number,
                ).value
            )

            if value:
                # Treat this first non-empty row as the header
                # only if it looks like a header.
                if normalize_match_key(value) in GENERIC_CATEGORY_HEADERS:
                    return row_number, column_number

                # Otherwise treat the first row as data.
                return 0, column_number

    raise ValueError(
        f"Could not locate a category column in worksheet "
        f"'{worksheet.title}'."
    )


def read_master_categories(
    workbook,
    sheet_index: int,
    category_type: str,
) -> Tuple[str, List[str]]:
    """
    Read authoritative master categories from a worksheet.

    Returns:

        worksheet title
        ordered unique category list
    """

    if sheet_index >= len(workbook.worksheets):
        raise ValueError(
            f"Research Funds.xlsx must contain at least "
            f"{sheet_index + 1} worksheets. "
            f"Found {len(workbook.worksheets)}."
        )

    worksheet = workbook.worksheets[sheet_index]

    header_row, category_column = locate_category_column(
        worksheet,
        category_type,
    )

    start_row = header_row + 1 if header_row > 0 else 1

    categories: List[str] = []
    seen = set()

    for row_number in range(
        start_row,
        worksheet.max_row + 1,
    ):
        value = clean_text(
            worksheet.cell(
                row=row_number,
                column=category_column,
            ).value
        )

        if not value:
            continue

        key = normalize_category_key(value)

        if not key:
            continue

        if key in GENERIC_CATEGORY_HEADERS:
            continue

        if key in seen:
            continue

        seen.add(key)
        categories.append(value)

    if not categories:
        raise ValueError(
            f"Worksheet '{worksheet.title}' contains no "
            f"{category_type} master categories."
        )

    return worksheet.title, categories


def read_all_research_data() -> Dict[str, Any]:
    """
    Read all Research Funds.xlsx information.

    Returns:

        {
          "funds": {...},
          "geographyWorksheet": "...",
          "sectorWorksheet": "...",
          "geographyCategories": [...],
          "sectorCategories": [...]
        }
    """

    workbook = open_excel_workbook(RESEARCH_EXCEL_FILE)

    try:
        research_funds = read_research_funds(workbook)

        geography_sheet, geography_categories = (
            read_master_categories(
                workbook,
                RESEARCH_GEOGRAPHY_SHEET_INDEX,
                "geography",
            )
        )

        sector_sheet, sector_categories = (
            read_master_categories(
                workbook,
                RESEARCH_SECTOR_SHEET_INDEX,
                "sector",
            )
        )

        return {
            "funds": research_funds,
            "geographyWorksheet": geography_sheet,
            "sectorWorksheet": sector_sheet,
            "geographyCategories": geography_categories,
            "sectorCategories": sector_categories,
        }

    finally:
        workbook.close()


# ============================================================
# HOLDINGS PIPELINE
# ============================================================

def extract_holdings_list(value: Any) -> List[Dict[str, Any]]:
    """
    Normalize a holdings list from the existing holdings pipeline.

    The function intentionally does not fabricate holdings.
    """

    if not isinstance(value, list):
        return []

    results: List[Dict[str, Any]] = []

    for holding in value:
        if not isinstance(holding, dict):
            continue

        name = clean_text(
            holding.get("name")
            or holding.get("holdingName")
            or holding.get("securityName")
        )

        if not name:
            continue

        record: Dict[str, Any] = {
            "rank": holding.get("rank"),
            "name": name,
            "weightPercent": holding.get("weightPercent"),
            "weightText": clean_text(
                holding.get("weightText")
            ),
        }

        results.append(record)

        if len(results) >= MAX_HOLDINGS:
            break

    return results


def load_holdings_file(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}

    try:
        payload = load_json(path)
    except Exception:
        return {}

    if isinstance(payload, dict):
        return payload

    return {}


def find_holdings_records(payload: Any) -> List[Dict[str, Any]]:
    """
    Recursively locate likely fund-level holdings records.

    This is deliberately tolerant of the existing holdings
    pipeline envelope.
    """

    found: List[Dict[str, Any]] = []

    def walk(value: Any) -> None:
        if isinstance(value, dict):

            if (
                "fundIdentifier" in value
                or "fundCode" in value
                or "fundName" in value
            ):
                if (
                    "holdings" in value
                    or "topHoldings" in value
                    or "count" in value
                    or "status" in value
                ):
                    found.append(value)

            for child in value.values():
                walk(child)

        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(payload)

    return found


def build_holdings_index() -> Dict[str, Dict[str, Any]]:
    """
    Load baseline and recovery outputs.

    Priority:

        baseline
        Recovery 1
        Recovery 2
        Recovery 3

    A later recovery result is only used when a fund is not
    already resolved by an earlier stage.

    No holdings are fabricated.
    """

    candidates: Dict[str, Dict[str, Any]] = {}

    source_files: List[Tuple[str, Path]] = [
        ("baseline", BASELINE_HOLDINGS_FILE),
    ]

    for index, directory in enumerate(RECOVERY_DIRECTORIES, start=1):

        possible_files = [
            directory / "all_holdings.json",
            directory / "holdings.json",
            directory / "recovered_holdings.json",
        ]

        for path in possible_files:
            if path.exists():
                source_files.append(
                    (f"recovery{index}", path)
                )

    for source_name, path in source_files:

        payload = load_holdings_file(path)

        if not payload:
            continue

        records = find_holdings_records(payload)

        for record in records:

            identifiers = [
                record.get("fundIdentifier"),
                record.get("fundCode"),
                record.get("fundName"),
            ]

            normalized_identifiers = [
                normalize_match_key(value)
                for value in identifiers
                if clean_text(value)
            ]

            holdings = (
                record.get("holdings")
                or record.get("topHoldings")
                or []
            )

            normalized_holdings = extract_holdings_list(
                holdings
            )

            if not normalized_holdings:
                continue

            normalized_record = dict(record)

            normalized_record["holdings"] = normalized_holdings
            normalized_record["_source"] = source_name

            for identifier in normalized_identifiers:
                if identifier not in candidates:
                    candidates[identifier] = normalized_record

    return candidates


def resolve_holdings(
    fund_identifier: str,
    fund_code: str,
    fund_name: str,
    holdings_index: Dict[str, Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    """
    Resolve holdings by the strongest available identity.

    Priority:

        fundIdentifier
        fundCode
        fundName
    """

    keys = [
        normalize_match_key(fund_identifier),
        normalize_match_key(fund_code),
        normalize_match_key(fund_name),
    ]

    for key in keys:
        if not key:
            continue

        record = holdings_index.get(key)

        if record:
            holdings = record.get("holdings") or []

            if holdings:
                return record

    return None


# ============================================================
# PRUACCESS / FUND DATA
# ============================================================

def recursively_find_fund_records(
    payload: Any,
) -> List[Dict[str, Any]]:
    """
    Recursively locate Prudential fund records.

    The existing extraction JSON can have different envelopes,
    so this does not assume one fixed top-level structure.
    """

    found: List[Dict[str, Any]] = []

    def looks_like_fund(record: Dict[str, Any]) -> bool:
        return bool(
            record.get("fundIdentifier")
            or record.get("fundName")
            or record.get("fundCode")
        )

    def walk(value: Any) -> None:

        if isinstance(value, dict):

            if looks_like_fund(value):
                found.append(value)

            for child in value.values():
                walk(child)

        elif isinstance(value, list):

            for child in value:
                walk(child)

    walk(payload)

    return found


def load_prudential_fund_records() -> Dict[str, Dict[str, Any]]:
    """
    Load available Prudential fund metadata from existing
    extraction outputs.

    Searches the known production locations.

    The first record for an identity is retained.
    """

    files: List[Path] = []

    if PRUACCESS_PRODUCTION_DIR.exists():
        files.extend(
            sorted(
                PRUACCESS_PRODUCTION_DIR.rglob("*.json")
            )
        )

    if PRUACCESS_FUNDS_DIR.exists():
        files.extend(
            sorted(
                PRUACCESS_FUNDS_DIR.rglob("*.json")
            )
        )

    index: Dict[str, Dict[str, Any]] = {}

    for path in files:

        try:
            payload = load_json(path)
        except Exception:
            continue

        records = recursively_find_fund_records(payload)

        for record in records:

            identifiers = [
                record.get("fundIdentifier"),
                record.get("fundCode"),
                record.get("fundName"),
            ]

            for identifier in identifiers:

                key = normalize_match_key(identifier)

                if not key:
                    continue

                if key not in index:
                    index[key] = record

    return index


def resolve_prudential_fund(
    fund_identifier: str,
    fund_code: str,
    fund_name: str,
    pruaccess_name: str,
    prudential_index: Dict[str, Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    """
    Resolve an upstream Prudential fund record.

    Priority follows the strongest identity available.
    """

    candidates = [
        fund_identifier,
        fund_code,
        fund_name,
        pruaccess_name,
    ]

    for candidate in candidates:

        key = normalize_match_key(candidate)

        if not key:
            continue

        record = prudential_index.get(key)

        if record:
            return record

    return None


# ============================================================
# FIELD EXTRACTION
# ============================================================

def first_nonempty(
    record: Dict[str, Any],
    *keys: str,
) -> Any:
    for key in keys:
        if key not in record:
            continue

        value = record.get(key)

        if not is_empty_value(value):
            return value

    return None


def extract_dividend_fields(
    source: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Extract dividend information.

    IMPORTANT:

        dividendUnit is the actual Prudential source field.

    hasDividend is derived from dividendRate.

    We deliberately do not calculate dividend units.
    """

    dividend_rate_raw = source.get("dividendRate")

    dividend_rate = clean_raw_text(
        dividend_rate_raw
    )

    dividend_unit = clean_text(
        source.get("dividendUnit")
    )

    return {
        "hasDividend": bool(dividend_rate),
        "dividendRate": dividend_rate or None,
        "dividendUnit": dividend_unit or None,
    }


def build_fund_metadata(
    source: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Build the public fund metadata block.

    Values are passed through from the Prudential extraction.
    """

    dividend = extract_dividend_fields(source)

    return {
        "fundCurrency": clean_text(
            source.get("fundCurrency")
        ),
        "unitCurrency": clean_text(
            source.get("unitCurrency")
        ),
        "assetClass": clean_text(
            source.get("assetClass")
        ),
        "assetSubClass": clean_text(
            source.get("assetSubClass")
        ),
        "riskClassification": clean_text(
            source.get("riskClassification")
        ),
        "bidPrice": clean_text(
            source.get("bidPrice")
        ),
        "offerPrice": clean_text(
            source.get("offerPrice")
        ),
        "valuationDate": clean_text(
            source.get("valuationDate")
        ),
        "inceptionDate": clean_text(
            source.get("inceptionDate")
        ),

        "cumulativeYtd": clean_text(
            source.get("cumulativeYtd")
        ),
        "cumulative1m": clean_text(
            source.get("cumulative1m")
        ),
        "cumulative3m": clean_text(
            source.get("cumulative3m")
        ),
        "cumulative6m": clean_text(
            source.get("cumulative6m")
        ),
        "cumulative1y": clean_text(
            source.get("cumulative1y")
        ),
        "cumulative3y": clean_text(
            source.get("cumulative3y")
        ),
        "cumulative5y": clean_text(
            source.get("cumulative5y")
        ),

        "annualised3y": clean_text(
            source.get("annualised3y")
        ),
        "annualised5y": clean_text(
            source.get("annualised5y")
        ),
        "annualised10y": clean_text(
            source.get("annualised10y")
        ),
        "annualisedSinceLaunch": clean_text(
            source.get("annualisedSinceLaunch")
        ),

        "factsheetUrl": clean_text(
            source.get("factsheetUrl")
        ),
        "prospectusUrl": clean_text(
            source.get("prospectusUrl")
        ),
        "productHighlightSheetUrl": clean_text(
            source.get("productHighlightSheetUrl")
        ),
        "annualReportUrl": clean_text(
            source.get("annualReportUrl")
        ),

        "fundObjective": clean_text(
            source.get("fundObjective")
        ),

        "investmentManager": clean_text(
            first_nonempty(
                source,
                "investmentManager",
                "fundManagers",
            )
        ),

        # Dividend fields.
        "hasDividend": dividend["hasDividend"],
        "dividendRate": dividend["dividendRate"],
        "dividendUnit": dividend["dividendUnit"],
    }


# ============================================================
# RESEARCH BLOCK
# ============================================================

def build_research_block(
    research_record: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Build the per-fund research block.

    Empty research remains explicitly empty.

    Research values are never inferred.
    """

    if not research_record:
        return {
            "geographic1": None,
            "geographic2": None,
            "sector1": None,
            "sector2": None,
        }

    return {
        "geographic1": (
            research_record.get("geographic1")
            or None
        ),
        "geographic2": (
            research_record.get("geographic2")
            or None
        ),
        "sector1": (
            research_record.get("sector1")
            or None
        ),
        "sector2": (
            research_record.get("sector2")
            or None
        ),
    }


# ============================================================
# FUND IDENTITY
# ============================================================

def get_fund_identifier(
    master_record: Dict[str, Any],
    source: Optional[Dict[str, Any]],
) -> str:
    """
    Resolve the Prudential fund identifier.

    Source identifier is preferred when available.
    """

    if source:
        identifier = clean_text(
            source.get("fundIdentifier")
        )

        if identifier:
            return identifier

    return ""


def get_fund_code(
    source: Optional[Dict[str, Any]],
) -> str:
    if not source:
        return ""

    return clean_text(
        source.get("fundCode")
    )


def get_fund_name(
    master_record: Dict[str, Any],
    source: Optional[Dict[str, Any]],
) -> str:
    """
    Master fund name / Prudential source fund name is preferred.

    If source has a fund name, use it.

    Otherwise fall back to PruAccess name.
    """

    if source:
        name = clean_text(
            source.get("fundName")
        )

        if name:
            return name

    return clean_text(
        master_record.get("pruAccessName")
    )


# ============================================================
# BID HISTORY
# ============================================================

def parse_date_value(value: Any) -> Optional[str]:
    """
    Normalize common date representations to YYYY-MM-DD.

    Existing historical BID dates are expected to already be
    clean. This is only a defensive normalizer.
    """

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    text = clean_text(value)

    if not text:
        return None

    # Already ISO.
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return text

    # Common DD/MM/YYYY.
    match = re.fullmatch(
        r"(\d{2})/(\d{2})/(\d{4})",
        text,
    )

    if match:
        day, month, year = match.groups()

        return f"{year}-{month}-{day}"

    return text


def parse_bid_value(value: Any) -> Optional[float]:
    """
    Convert BID value into a number where possible.

    Currency symbols and commas are removed.
    """

    if value is None:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)

    if not text:
        return None

    text = text.replace(",", "")
    text = text.replace("$", "")
    text = text.replace("SGD", "")
    text = text.strip()

    try:
        return float(text)
    except ValueError:
        return None


def recursively_find_bid_observations(
    payload: Any,
) -> List[Dict[str, Any]]:
    """
    Recursively locate historical BID observation records.

    Accepted common structures include:

        {"date": "...", "bid": "..."}
        {"valuationDate": "...", "bidPrice": "..."}
        {"date": "...", "bidPrice": "..."}
    """

    observations: List[Dict[str, Any]] = []

    def walk(value: Any) -> None:

        if isinstance(value, dict):

            date_value = first_nonempty(
                value,
                "date",
                "valuationDate",
                "bidDate",
            )

            bid_value = first_nonempty(
                value,
                "bid",
                "bidPrice",
                "value",
                "price",
            )

            if (
                date_value is not None
                and bid_value is not None
            ):
                parsed_date = parse_date_value(
                    date_value
                )

                parsed_bid = parse_bid_value(
                    bid_value
                )

                if (
                    parsed_date is not None
                    and parsed_bid is not None
                ):
                    observations.append(
                        {
                            "date": parsed_date,
                            "bid": parsed_bid,
                        }
                    )

            for child in value.values():
                walk(child)

        elif isinstance(value, list):

            for child in value:
                walk(child)

    walk(payload)

    return observations


def deduplicate_bid_observations(
    observations: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Deduplicate by date.

    The raw historical BID dataset should contain only one
    observation for a given date.
    """

    by_date: Dict[str, Dict[str, Any]] = {}

    for observation in observations:

        date_value = observation.get("date")
        bid_value = observation.get("bid")

        if not date_value:
            continue

        if bid_value is None:
            continue

        by_date[date_value] = {
            "date": date_value,
            "bid": bid_value,
        }

    return [
        by_date[key]
        for key in sorted(by_date.keys())
    ]


def load_bid_history_for_identifier(
    fund_identifier: str,
    fund_code: str,
    fund_name: str,
) -> List[Dict[str, Any]]:
    """
    Search existing PruAccess output for historical BID data.

    This function does not create or carry forward missing raw
    observations.
    """

    search_keys = [
        normalize_match_key(fund_identifier),
        normalize_match_key(fund_code),
        normalize_match_key(fund_name),
    ]

    search_keys = [
        key for key in search_keys if key
    ]

    if not search_keys:
        return []

    directories = [
        PRUACCESS_PRODUCTION_DIR,
        PRUACCESS_FUNDS_DIR,
    ]

    observations: List[Dict[str, Any]] = []

    for directory in directories:

        if not directory.exists():
            continue

        for path in directory.rglob("*.json"):

            # Avoid reading the output files produced by this
            # build script if they happen to be under data/.
            if DATA_DIR in path.parents:
                continue

            try:
                payload = load_json(path)
            except Exception:
                continue

            text_blob = ""

            if isinstance(payload, dict):
                identity_values = [
                    payload.get("fundIdentifier"),
                    payload.get("fundCode"),
                    payload.get("fundName"),
                    payload.get("pruAccessName"),
                ]

                identity_keys = {
                    normalize_match_key(value)
                    for value in identity_values
                    if clean_text(value)
                }

                if identity_keys.intersection(search_keys):
                    observations.extend(
                        recursively_find_bid_observations(
                            payload
                        )
                    )

            elif isinstance(payload, list):
                records = recursively_find_fund_records(
                    payload
                )

                for record in records:

                    identity_values = [
                        record.get("fundIdentifier"),
                        record.get("fundCode"),
                        record.get("fundName"),
                        record.get("pruAccessName"),
                    ]

                    identity_keys = {
                        normalize_match_key(value)
                        for value in identity_values
                        if clean_text(value)
                    }

                    if identity_keys.intersection(
                        search_keys
                    ):
                        observations.extend(
                            recursively_find_bid_observations(
                                record
                            )
                        )

    return deduplicate_bid_observations(
        observations
    )


# ============================================================
# MAIN BUILD
# ============================================================

def main() -> int:

    print("VGrat FMS - Build Funds Data")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. Read controlling master universe.
    # --------------------------------------------------------

    print("Reading master fund universe...")

    master_funds = read_master_funds()

    if not master_funds:
        raise RuntimeError(
            "No funds were found in Funds Links.xlsm."
        )

    print(
        f"Master universe: {len(master_funds)} funds"
    )

    # --------------------------------------------------------
    # 2. Read Research Funds.xlsx.
    # --------------------------------------------------------

    print("Reading Research Funds.xlsx...")

    research_data = read_all_research_data()

    research_funds = research_data["funds"]

    geography_categories = research_data[
        "geographyCategories"
    ]

    sector_categories = research_data[
        "sectorCategories"
    ]

    print(
        f"Research records: {len(research_funds)}"
    )

    print(
        f"Geography master categories: "
        f"{len(geography_categories)}"
    )

    print(
        f"Sector master categories: "
        f"{len(sector_categories)}"
    )

    # --------------------------------------------------------
    # 3. Load holdings.
    # --------------------------------------------------------

    print("Loading holdings pipeline outputs...")

    holdings_index = build_holdings_index()

    print(
        f"Holdings identity records: "
        f"{len(holdings_index)}"
    )

    # --------------------------------------------------------
    # 4. Load Prudential fund extraction.
    # --------------------------------------------------------

    print("Loading Prudential fund extraction...")

    prudential_index = (
        load_prudential_fund_records()
    )

    print(
        f"Prudential fund records indexed: "
        f"{len(prudential_index)}"
    )

    # --------------------------------------------------------
    # 5. Build funds.
    # --------------------------------------------------------

    funds: List[Dict[str, Any]] = []

    research_matched = 0
    research_missing = 0

    holdings_matched = 0
    holdings_missing = 0

    prudential_matched = 0
    prudential_missing = 0

    dividend_count = 0

    unresolved: List[Dict[str, Any]] = []

    for master_record in master_funds:

        prudential_url = clean_text(
            master_record.get("prudentialUrl")
        )

        pruaccess_name = clean_text(
            master_record.get("pruAccessName")
        )

        # ----------------------------------------------------
        # Resolve Prudential source.
        # ----------------------------------------------------

        source = resolve_prudential_fund(
            fund_identifier="",
            fund_code="",
            fund_name="",
            pruaccess_name=pruaccess_name,
            prudential_index=prudential_index,
        )

        if source:
            prudential_matched += 1
        else:
            prudential_missing += 1

        fund_identifier = get_fund_identifier(
            master_record,
            source,
        )

        fund_code = get_fund_code(
            source,
        )

        fund_name = get_fund_name(
            master_record,
            source,
        )

        # ----------------------------------------------------
        # Research lookup.
        #
        # Match against the researched fund name.
        # ----------------------------------------------------

        research_record = research_funds.get(
            normalize_match_key(fund_name)
        )

        if research_record is None:
            research_record = research_funds.get(
                normalize_match_key(pruaccess_name)
            )

        if research_record:
            research_matched += 1
        else:
            research_missing += 1

        # ----------------------------------------------------
        # Holdings lookup.
        # ----------------------------------------------------

        holdings_record = resolve_holdings(
            fund_identifier=fund_identifier,
            fund_code=fund_code,
            fund_name=fund_name,
            holdings_index=holdings_index,
        )

        if holdings_record:
            holdings_matched += 1
        else:
            holdings_missing += 1

        # ----------------------------------------------------
        # Fund metadata.
        # ----------------------------------------------------

        if source:
            fund_metadata = build_fund_metadata(
                source
            )
        else:
            # Keep a stable schema even when the source
            # extraction cannot be resolved.
            fund_metadata = {
                "fundCurrency": "",
                "unitCurrency": "",
                "assetClass": "",
                "assetSubClass": "",
                "riskClassification": "",
                "bidPrice": "",
                "offerPrice": "",
                "valuationDate": "",
                "inceptionDate": "",
                "cumulativeYtd": "",
                "cumulative1m": "",
                "cumulative3m": "",
                "cumulative6m": "",
                "cumulative1y": "",
                "cumulative3y": "",
                "cumulative5y": "",
                "annualised3y": "",
                "annualised5y": "",
                "annualised10y": "",
                "annualisedSinceLaunch": "",
                "factsheetUrl": "",
                "prospectusUrl": "",
                "productHighlightSheetUrl": "",
                "annualReportUrl": "",
                "fundObjective": "",
                "investmentManager": "",
                "hasDividend": False,
                "dividendRate": None,
                "dividendUnit": None,
            }

        if fund_metadata.get("hasDividend"):
            dividend_count += 1

        # ----------------------------------------------------
        # Top holdings.
        # ----------------------------------------------------

        if holdings_record:

            source_holdings = (
                holdings_record.get("holdings")
                or []
            )

            top_holdings = {
                "status": clean_text(
                    holdings_record.get("status")
                ) or "published",
                "source": clean_text(
                    holdings_record.get("_source")
                ),
                "parser": clean_text(
                    holdings_record.get("parser")
                ),
                "factsheetUrl": clean_text(
                    holdings_record.get(
                        "factsheetUrl"
                    )
                ),
                "factsheetDocumentDate": clean_text(
                    holdings_record.get(
                        "factsheetDocumentDate"
                    )
                ),
                "factsheetDataAsAt": clean_text(
                    holdings_record.get(
                        "factsheetDataAsAt"
                    )
                ),
                "count": len(source_holdings),
                "holdings": source_holdings,
            }

        else:

            top_holdings = {
                "status": "unavailable",
                "source": "",
                "parser": "",
                "factsheetUrl": fund_metadata.get(
                    "factsheetUrl",
                    "",
                ),
                "factsheetDocumentDate": "",
                "factsheetDataAsAt": "",
                "count": 0,
                "holdings": [],
            }

        # ----------------------------------------------------
        # Research.
        # ----------------------------------------------------

        research_block = build_research_block(
            research_record
        )

        # ----------------------------------------------------
        # Final fund record.
        # ----------------------------------------------------

        fund_record = {
            "excelRow": master_record.get(
                "excelRow"
            ),
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "fundName": fund_name,
            "pruAccessName": pruaccess_name,
            "prudentialUrl": prudential_url,

            "fund": fund_metadata,

            "research": research_block,

            "topHoldings": top_holdings,
        }

        funds.append(fund_record)

        # ----------------------------------------------------
        # Unresolved tracking.
        # ----------------------------------------------------

        issues = []

        if not source:
            issues.append("prudential_source_missing")

        if not research_record:
            issues.append("research_missing")

        if not holdings_record:
            issues.append("holdings_missing")

        if issues:
            unresolved.append(
                {
                    "excelRow": master_record.get(
                        "excelRow"
                    ),
                    "fundIdentifier": fund_identifier,
                    "fundCode": fund_code,
                    "fundName": fund_name,
                    "pruAccessName": pruaccess_name,
                    "issues": issues,
                }
            )

    # --------------------------------------------------------
    # 6. Build BID history.
    # --------------------------------------------------------

    print("Building historical BID data...")

    bid_history_funds: List[Dict[str, Any]] = []

    bid_observation_count = 0
    bid_funds_with_history = 0

    for fund in funds:

        fund_identifier = clean_text(
            fund.get("fundIdentifier")
        )

        fund_code = clean_text(
            fund.get("fundCode")
        )

        fund_name = clean_text(
            fund.get("fundName")
        )

        observations = load_bid_history_for_identifier(
            fund_identifier=fund_identifier,
            fund_code=fund_code,
            fund_name=fund_name,
        )

        if observations:
            bid_funds_with_history += 1

        bid_observation_count += len(
            observations
        )

        bid_history_funds.append(
            {
                "excelRow": fund.get("excelRow"),
                "fundIdentifier": fund_identifier,
                "fundCode": fund_code,
                "fundName": fund_name,
                "pruAccessName": clean_text(
                    fund.get("pruAccessName")
                ),
                "observations": observations,
            }
        )

    # --------------------------------------------------------
    # 7. Build funds.json.
    # --------------------------------------------------------

    generated_at = utc_now_iso()

    funds_payload = {
        "schemaVersion": SCHEMA_VERSION,
        "generatedAt": generated_at,

        "research": {
            "source": RESEARCH_EXCEL_FILE.name,

            "fundResearchWorksheet": (
                RESEARCH_FUND_SHEET_NAME
            ),

            "masterCategoryWorksheets": {
                "geography": research_data[
                    "geographyWorksheet"
                ],
                "sector": research_data[
                    "sectorWorksheet"
                ],
            },

            # IMPORTANT:
            # These are authoritative UI filter values.
            #
            # They are copied directly from worksheets 2 and 3.
            #
            # They are NOT derived from the fund records.
            "masterCategories": {
                "geography": geography_categories,
                "sector": sector_categories,
            },
        },

        "summary": {
            "fundCount": len(funds),

            "researchMatched": research_matched,
            "researchMissing": research_missing,

            "holdingsMatched": holdings_matched,
            "holdingsMissing": holdings_missing,

            "prudentialMatched": prudential_matched,
            "prudentialMissing": prudential_missing,

            "dividendFunds": dividend_count,

            "unresolvedCount": len(unresolved),
        },

        "funds": funds,
    }

    # --------------------------------------------------------
    # 8. Build bid_history.json.
    # --------------------------------------------------------

    bid_history_payload = {
        "schemaVersion": SCHEMA_VERSION,
        "generatedAt": generated_at,

        "summary": {
            "fundCount": len(
                bid_history_funds
            ),
            "fundsWithHistory": (
                bid_funds_with_history
            ),
            "observationCount": (
                bid_observation_count
            ),
        },

        "funds": bid_history_funds,
    }

    # --------------------------------------------------------
    # 9. Write outputs.
    # --------------------------------------------------------

    print("Writing data/funds.json...")

    write_json_atomic(
        FUNDS_OUTPUT_FILE,
        funds_payload,
    )

    print("Writing data/bid_history.json...")

    write_json_atomic(
        BID_HISTORY_OUTPUT_FILE,
        bid_history_payload,
    )

    # --------------------------------------------------------
    # 10. Build validation.
    # --------------------------------------------------------

    if not ALLOW_UNRESOLVED and unresolved:
        print()
        print(
            f"WARNING: {len(unresolved)} funds have "
            f"unresolved source/research/holdings data."
        )

        for item in unresolved[:20]:
            print(
                f"  - {item['fundName']} "
                f"({', '.join(item['issues'])})"
            )

        if len(unresolved) > 20:
            print(
                f"  ... and "
                f"{len(unresolved) - 20} more"
            )

    # --------------------------------------------------------
    # 11. Final checks.
    # --------------------------------------------------------

    if len(funds) != len(master_funds):
        raise RuntimeError(
            "Fatal: output fund count does not match "
            "the controlling Funds Links.xlsm universe."
        )

    if not geography_categories:
        raise RuntimeError(
            "Fatal: Geography master categories are empty."
        )

    if not sector_categories:
        raise RuntimeError(
            "Fatal: Sector master categories are empty."
        )

    print()
    print("Build complete.")
    print(
        f"Funds: {len(funds)}"
    )
    print(
        f"Dividend funds: {dividend_count}"
    )
    print(
        f"Geography filter categories: "
        f"{len(geography_categories)}"
    )
    print(
        f"Sector filter categories: "
        f"{len(sector_categories)}"
    )
    print(
        f"BID observations: "
        f"{bid_observation_count}"
    )
    print()
    print(
        f"Output: {FUNDS_OUTPUT_FILE}"
    )
    print(
        f"Output: {BID_HISTORY_OUTPUT_FILE}"
    )

    return 0


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    try:
        sys.exit(main())

    except KeyboardInterrupt:
        print(
            "\nBuild cancelled.",
            file=sys.stderr,
        )
        sys.exit(130)

    except Exception as exc:
        print(
            f"\nERROR: {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
