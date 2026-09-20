#!/usr/bin/env python3

"""
PruAccess extraction validator.

PURPOSE
=======

Validate the output produced by:

    scripts/test_pruaccess.py

MASTER SOURCE
=============

Funds Links.xlsm

Column A:
    Prudential fund URL

The Excel file defines the complete fund universe.

IMPORTANT OUTPUT CONTRACT
=========================

The extractor does NOT create a separate pagination.json file.

Pagination diagnostics are stored inside:

    funds/<excelRow>_<fundIdentifier>/window_results.json

Each historical window contains its own pagination list.

This validator therefore validates pagination directly from:

    window_results.json

HARD RULES
==========

1. Excel Column A defines the master fund universe.
2. Every Excel-master fund must have a successful output.
3. PruAccess must be historical BID data.
4. Every historical observation must contain a valid date.
5. Every historical observation must contain a valid numeric BID price.
6. No duplicate historical dates are allowed.
7. Historical observations must be chronologically ordered.
8. The oldest observation must be:
       - on the Prudential inception date, OR
       - no more than 7 calendar days after inception.
9. No observation may be before the Prudential inception date.
10. No observation may be after the requested/current end date.
11. Consecutive actual historical observations may have gaps,
    but each gap must be no more than 7 calendar days.
12. Date gaps must NEVER be filled.
13. No interpolation is permitted.
14. No estimation is permitted.
15. No synthetic observations are permitted.
16. Every historical window must exist.
17. Every historical window must pass validation.
18. Every expected pagination page must exist.
19. Pagination pages must be continuous.
20. Every pagination page must have status=success.
21. Pagination row counts must match extracted observations.
22. Fund-level historical observation totals must reconcile.
23. Fund-level page totals must reconcile.
24. Fund-level window totals must reconcile.
25. Consolidated all_bid_history.json must reconcile with individual funds.
26. run_summary.json must report the complete Excel universe.
27. validation.json must report the complete Excel universe.
28. No silent partial extraction is accepted.

NO SYNTHETIC DATA
=================

This validator does not create, repair, interpolate, estimate,
carry-forward, or fabricate any historical observations.

It only validates data already extracted by test_pruaccess.py.
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR
    / "funds"
)

RUN_SUMMARY_FILE = (
    OUTPUT_DIR
    / "run_summary.json"
)

ALL_FUNDS_FILE = (
    OUTPUT_DIR
    / "all_funds.json"
)

ALL_BID_HISTORY_FILE = (
    OUTPUT_DIR
    / "all_bid_history.json"
)

VALIDATION_FILE = (
    OUTPUT_DIR
    / "validation.json"
)

PAGE_SIZE = 20

# ------------------------------------------------------------
# HARD HISTORICAL DATE GAP RULE
#
# Consecutive actual PruAccess observations may be separated
# by 1-7 calendar days.
#
# A gap greater than 7 calendar days fails validation.
# ------------------------------------------------------------

MAX_BID_DATE_GAP_DAYS = 7

# ------------------------------------------------------------
# Numerical BID comparison tolerance.
#
# Used when comparing current Prudential BID and latest
# PruAccess BID where both values are available.
# ------------------------------------------------------------

PRICE_TOLERANCE = 0.0002

# Maximum number of examples printed for an error category.
MAX_ERROR_EXAMPLES = 20


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(
    value: Any,
) -> str:
    """
    Normalize a value into clean single-spaced text.
    """

    if value is None:
        return ""

    return " ".join(
        str(value)
        .replace("\xa0", " ")
        .split()
    ).strip()


def normalize_text(
    value: Any,
) -> str:
    """
    Normalize text for exact identity comparisons where
    whitespace/case should not create false differences.
    """

    return clean_text(
        value
    ).casefold()


def load_json(
    path: Path,
) -> Any:
    """
    Load JSON from disk.
    """

    if not path.exists():

        raise FileNotFoundError(
            f"Required JSON file not found: {path}"
        )

    return json.loads(
        path.read_text(
            encoding="utf-8"
        )
    )


def parse_date(
    value: Any,
) -> date:
    """
    Parse supported date representations.

    Supports:
        DD-Mon-YYYY
        DD-Month-YYYY
        DD/MM/YYYY
        YYYY-MM-DD
        YYYY/MM/DD
    """

    if isinstance(
        value,
        datetime,
    ):

        return value.date()

    if isinstance(
        value,
        date,
    ):

        return value

    text = clean_text(
        value
    )

    formats = [
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d/%b/%Y",
        "%d/%B/%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d %b %Y",
        "%d %B %Y",
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                text,
                fmt,
            ).date()

        except ValueError:

            continue

    raise ValueError(
        f"Unable to parse date: {value!r}"
    )


def format_date(
    value: date,
) -> str:
    """
    Format date consistently.
    """

    return value.strftime(
        "%d-%b-%Y"
    )


def parse_numeric_price(
    value: Any,
) -> float:
    """
    Parse a historical BID price.

    Only actual numeric values are accepted.

    Currency symbols and commas are removed because the
    extractor may preserve display formatting in some fields.

    Empty / '-' / unavailable values are invalid for a
    historical BID observation.
    """

    if value is None:

        raise ValueError(
            "BID price is missing."
        )

    text = clean_text(
        value
    )

    if not text:

        raise ValueError(
            "BID price is empty."
        )

    if text in {
        "-",
        "—",
        "N/A",
        "NA",
        "null",
        "None",
    }:

        raise ValueError(
            f"BID price is unavailable: {value!r}"
        )

    cleaned = (
        text
        .replace(",", "")
        .replace("$", "")
        .replace("S$", "")
        .replace("SGD", "")
        .strip()
    )

    try:

        numeric = float(
            cleaned
        )

    except ValueError as error:

        raise ValueError(
            f"BID price is not numeric: {value!r}"
        ) from error

    if numeric <= 0:

        raise ValueError(
            f"BID price must be greater than zero: {value!r}"
        )

    return numeric


def values_close(
    first: float,
    second: float,
    tolerance: float = PRICE_TOLERANCE,
) -> bool:
    """
    Compare two prices using absolute tolerance.
    """

    return abs(
        first - second
    ) <= tolerance


# ============================================================
# EXCEL MASTER UNIVERSE
# ============================================================

def read_excel_funds() -> list[dict[str, Any]]:
    """
    Read every populated URL from Column A.

    Column B is retained as reference metadata only.

    The number of funds is never hardcoded.
    """

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel master file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    try:

        worksheet = workbook[
            workbook.sheetnames[0]
        ]

        funds = []

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):

            prudential_url = clean_text(
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

            if not prudential_url:

                continue

            funds.append(
                {
                    "excelRow":
                        row_number,

                    "prudentialUrl":
                        prudential_url,

                    "pruAccessName":
                        pruaccess_name,
                }
            )

    finally:

        workbook.close()

    if not funds:

        raise ValueError(
            "No populated Prudential fund URLs were found "
            "in Column A of Funds Links.xlsm."
        )

    return funds


# ============================================================
# FUND DIRECTORY DISCOVERY
# ============================================================

def get_fund_directory(
    excel_row: int,
) -> Path | None:
    """
    Locate the successful fund directory for an Excel row.

    Successful directories are created by test_pruaccess.py
    using:

        <excelRow>_<fundIdentifier>

    Failed directories use:

        <excelRow>_failed

    Failed directories are explicitly excluded.
    """

    if not FUNDS_OUTPUT_DIR.exists():

        return None

    prefix = (
        f"{excel_row}_"
    )

    candidates = [
        path
        for path in FUNDS_OUTPUT_DIR.iterdir()
        if (
            path.is_dir()
            and path.name.startswith(prefix)
            and "failed" not in path.name.lower()
        )
    ]

    if not candidates:

        return None

    if len(candidates) > 1:

        raise ValueError(
            f"Multiple successful fund directories found "
            f"for Excel row {excel_row}: "
            f"{[path.name for path in candidates]}"
        )

    return candidates[0]


# ============================================================
# REQUIRED FILE VALIDATION
# ============================================================

def validate_required_files(
    fund_directory: Path,
) -> list[str]:
    """
    Validate the files actually produced by the extractor.

    IMPORTANT:
    pagination.json is intentionally NOT required.

    Pagination is stored in window_results.json.
    """

    required_files = [
        "summary.json",
        "bid_history.json",
        "date_windows.json",
        "window_diagnostics.json",
        "window_results.json",
        "inception_validation.json",
        "prudential_fund.json",
    ]

    errors = []

    for filename in required_files:

        path = (
            fund_directory
            / filename
        )

        if not path.exists():

            errors.append(
                f"Missing required file: {filename}"
            )

    return errors


# ============================================================
# EXTRACTOR OUTPUT STRUCTURE VALIDATION
# ============================================================

def validate_summary_structure(
    summary: Any,
    excel_fund: dict[str, Any],
) -> list[str]:

    errors = []

    if not isinstance(
        summary,
        dict,
    ):

        return [
            "summary.json must contain a JSON object."
        ]

    if summary.get("status") != "success":

        errors.append(
            "summary.json status is not 'success'."
        )

    if summary.get("excelRow") != excel_fund["excelRow"]:

        errors.append(
            "summary.json Excel row does not match "
            "the Excel master row."
        )

    summary_url = normalize_text(
        summary.get(
            "prudentialUrl"
        )
    )

    excel_url = normalize_text(
        excel_fund.get(
            "prudentialUrl"
        )
    )

    if summary_url != excel_url:

        errors.append(
            "summary.json Prudential URL does not match "
            "the Excel master URL."
        )

    historical_price_type = clean_text(
        summary.get(
            "historicalPriceType"
        )
    ).upper()

    if historical_price_type != "BID":

        errors.append(
            "summary.json historicalPriceType is not BID."
        )

    if "pagination" in summary:

        errors.append(
            "summary.json contains an unexpected top-level "
            "pagination field; pagination should be stored "
            "inside window_results.json."
        )

    return errors


# ============================================================
# PRUDENTIAL FUND VALIDATION
# ============================================================

def validate_prudential_fund(
    prudential: Any,
    excel_fund: dict[str, Any],
) -> list[str]:

    errors = []

    if not isinstance(
        prudential,
        dict,
    ):

        return [
            "prudential_fund.json must contain a JSON object."
        ]

    fund_name = clean_text(
        prudential.get(
            "fundName"
        )
    )

    fund_identifier = clean_text(
        prudential.get(
            "fundIdentifier"
        )
    )

    inception_raw = clean_text(
        prudential.get(
            "inceptionDate"
        )
    )

    if not fund_name:

        errors.append(
            "Prudential fund name is missing."
        )

    if not fund_identifier:

        errors.append(
            "Prudential fundIdentifier is missing."
        )

    if not inception_raw:

        errors.append(
            "Prudential inceptionDate is missing."
        )

    if inception_raw:

        try:

            parse_date(
                inception_raw
            )

        except ValueError as error:

            errors.append(
                f"Invalid Prudential inceptionDate: {error}"
            )

    current_bid = prudential.get(
        "bidPrice"
    )

    if current_bid in (
        None,
        "",
        "-",
    ):

        errors.append(
            "Prudential current BID price is missing."
        )

    else:

        try:

            parse_numeric_price(
                current_bid
            )

        except ValueError as error:

            errors.append(
                f"Invalid Prudential current BID price: {error}"
            )

    return errors


# ============================================================
# BID HISTORY STRUCTURE
# ============================================================

def validate_bid_history_structure(
    bid_history: Any,
) -> list[str]:

    errors = []

    if not isinstance(
        bid_history,
        dict,
    ):

        return [
            "bid_history.json must contain a JSON object."
        ]

    source = clean_text(
        bid_history.get(
            "source"
        )
    )

    if source != "PruAccess":

        errors.append(
            "bid_history.json source is not PruAccess."
        )

    price_type = clean_text(
        bid_history.get(
            "priceType"
        )
    ).upper()

    if price_type != "BID":

        errors.append(
            "bid_history.json priceType is not BID."
        )

    observations = bid_history.get(
        "observations"
    )

    if not isinstance(
        observations,
        list,
    ):

        errors.append(
            "bid_history.json observations must be a list."
        )

    else:

        if not observations:

            errors.append(
                "bid_history.json contains zero observations."
            )

        observation_count = bid_history.get(
            "observationCount"
        )

        if not isinstance(
            observation_count,
            int,
        ):

            errors.append(
                "bid_history.json observationCount "
                "must be an integer."
            )

        elif observation_count != len(
            observations
        ):

            errors.append(
                "bid_history.json observationCount does not "
                "match the actual observation list length."
            )

    return errors


# ============================================================
# OBSERVATION VALIDATION
# ============================================================

def validate_observations(
    observations: list[dict[str, Any]],
) -> dict[str, Any]:
    """
    Validate every historical BID observation.

    Rules:
        - valid date
        - valid numeric positive BID
        - no duplicate dates
        - chronological ordering
        - no synthetic marker fields
        - consecutive gap <= 7 calendar days
    """

    errors = []
    warnings = []

    if not observations:

        errors.append(
            "Historical observation list is empty."
        )

        return {
            "errors": errors,
            "warnings": warnings,
            "parsedDates": [],
            "oldestDate": None,
            "newestDate": None,
            "dateGaps": [],
            "duplicateDates": [],
        }

    parsed_dates = []

    duplicate_map: dict[str, list[dict[str, Any]]] = {}

    suspicious_keys = [
        "synthetic",
        "estimated",
        "estimate",
        "interpolated",
        "interpolation",
        "fabricated",
        "simulated",
    ]

    for index, observation in enumerate(
        observations,
        start=1,
    ):

        if not isinstance(
            observation,
            dict,
        ):

            errors.append(
                f"Observation {index} is not a JSON object."
            )

            continue

        if "date" not in observation:

            errors.append(
                f"Observation {index} is missing date."
            )

        else:

            try:

                observation_date = parse_date(
                    observation["date"]
                )

                parsed_dates.append(
                    observation_date
                )

                date_key = (
                    observation_date.isoformat()
                )

                duplicate_map.setdefault(
                    date_key,
                    [],
                ).append(
                    observation
                )

            except ValueError as error:

                errors.append(
                    f"Observation {index} has invalid date: {error}"
                )

        if "bidPrice" not in observation:

            errors.append(
                f"Observation {index} is missing bidPrice."

            )

        else:

            try:

                parse_numeric_price(
                    observation["bidPrice"]
                )

            except ValueError as error:

                errors.append(
                    f"Observation {index} has invalid bidPrice: "
                    f"{error}"
                )

        for key in observation.keys():

            key_lower = str(
                key
            ).casefold()

            for suspicious in suspicious_keys:

                if suspicious in key_lower:

                    errors.append(
                        "Historical observation contains "
                        f"prohibited synthetic-data field "
                        f"'{key}' at observation {index}."
                    )

                    break

    duplicate_dates = []

    for date_key, rows in duplicate_map.items():

        if len(rows) > 1:

            duplicate_dates.append(
                {
                    "date":
                        date_key,

                    "count":
                        len(rows),
                }
            )

    if duplicate_dates:

        errors.append(
            "Duplicate historical BID dates detected: "
            f"{len(duplicate_dates)} duplicate date(s)."
        )

    # --------------------------------------------------------
    # Chronological order
    # --------------------------------------------------------

    for index in range(
        1,
        len(parsed_dates),
    ):

        previous_date = (
            parsed_dates[index - 1]
        )

        current_date = (
            parsed_dates[index]
        )

        if current_date <= previous_date:

            if current_date == previous_date:

                errors.append(
                    "Historical observations are not strictly "
                    "chronological because a duplicate date exists."
                )

            else:

                errors.append(
                    "Historical observations are not in "
                    "ascending chronological order."
                )

            break

    # --------------------------------------------------------
    # Consecutive actual observation gap
    #
    # Maximum allowed = 7 calendar days.
    #
    # Example:
    #
    # 10-Sep -> 11-Sep = 1 day
    # 10-Sep -> 17-Sep = 7 days = allowed
    # 10-Sep -> 18-Sep = 8 days = FAIL
    #
    # We do NOT fill these gaps.
    # --------------------------------------------------------

    excessive_gaps = []
    date_gaps = []

    for index in range(
        1,
        len(parsed_dates),
    ):

        previous_date = (
            parsed_dates[index - 1]
        )

        current_date = (
            parsed_dates[index]
        )

        gap_days = (
            current_date
            - previous_date
        ).days

        if gap_days > 1:

            date_gaps.append(
                {
                    "fromDate":
                        previous_date.isoformat(),

                    "toDate":
                        current_date.isoformat(),

                    "calendarDays":
                        gap_days,

                    "missingCalendarDays":
                        gap_days - 1,
                }
            )

        if gap_days > MAX_BID_DATE_GAP_DAYS:

            excessive_gaps.append(
                {
                    "fromDate":
                        previous_date.isoformat(),

                    "toDate":
                        current_date.isoformat(),

                    "calendarDays":
                        gap_days,
                }
            )

    if excessive_gaps:

        errors.append(
            "Historical BID date gap exceeds the maximum "
            f"allowed {MAX_BID_DATE_GAP_DAYS} calendar days. "
            f"Found {len(excessive_gaps)} excessive gap(s)."
        )

    oldest_date = (
        min(parsed_dates)
        if parsed_dates
        else None
    )

    newest_date = (
        max(parsed_dates)
        if parsed_dates
        else None
    )

    return {
        "errors":
            errors,

        "warnings":
            warnings,

        "parsedDates":
            parsed_dates,

        "oldestDate":
            oldest_date,

        "newestDate":
            newest_date,

        "dateGaps":
            date_gaps,

        "excessiveGaps":
            excessive_gaps,

        "duplicateDates":
            duplicate_dates,
    }


# ============================================================
# WINDOW VALIDATION
# ============================================================

def validate_date_windows(
    date_windows: Any,
    window_results: Any,
) -> tuple[
    list[str],
    list[dict[str, Any]],
]:
    """
    Validate the historical date-window structure.

    Each window must:
        - be numbered continuously
        - have start/end dates
        - not exceed the 10-year window rule
        - not overlap another window
        - not contain gaps between windows
        - correspond to a window result
    """

    errors = []

    normalized_windows = []

    if not isinstance(
        date_windows,
        list,
    ):

        return [
            "date_windows.json must contain a list."
        ], []

    if not date_windows:

        return [
            "date_windows.json contains zero windows."
        ], []

    for index, window in enumerate(
        date_windows,
        start=1,
    ):

        if not isinstance(
            window,
            dict,
        ):

            errors.append(
                f"Date window {index} is not a JSON object."
            )

            continue

        window_number = window.get(
            "windowNumber"
        )

        if not isinstance(
            window_number,
            int,
        ):

            errors.append(
                f"Date window {index} has invalid windowNumber."
            )

            continue

        if window_number != index:

            errors.append(
                "Date windows are not numbered continuously. "
                f"Expected {index}, got {window_number}."
            )

        try:

            start_date = parse_date(
                window.get(
                    "startDate"
                )
            )

            end_date = parse_date(
                window.get(
                    "endDate"
                )

            )

        except Exception as error:

            errors.append(
                f"Window {window_number} has invalid dates: "
                f"{error}"
            )

            continue

        if end_date < start_date:

            errors.append(
                f"Window {window_number} end date is before "
                "its start date."
            )

        # ----------------------------------------------------
        # 10-year maximum.
        #
        # Extractor uses:
        #
        # add_calendar_years(start, 10) - 1 day
        #
        # Therefore end date must be strictly before:
        #
        # start + 10 years
        # ----------------------------------------------------

        try:

            ten_year_limit = start_date.replace(
                year=start_date.year + 10
            )

        except ValueError:

            ten_year_limit = start_date.replace(
                year=start_date.year + 10,
                day=28,
            )

        if end_date >= ten_year_limit:

            errors.append(
                f"Window {window_number} exceeds the "
                "maximum 10-year PruAccess window."
            )

        normalized_windows.append(
            {
                "windowNumber":
                    window_number,

                "startDate":
                    start_date,

                "endDate":
                    end_date,
            }
        )

    # --------------------------------------------------------
    # Window continuity
    # --------------------------------------------------------

    normalized_windows.sort(
        key=lambda item:
            item["windowNumber"]
    )

    for index in range(
        1,
        len(normalized_windows),
    ):

        previous = normalized_windows[
            index - 1
        ]

        current = normalized_windows[
            index
        ]

        expected_start = (
            previous["endDate"]
            + timedelta(days=1)
        )

        if current["startDate"] != expected_start:

            errors.append(
                "Historical windows contain a gap or overlap "
                f"between window "
                f"{previous['windowNumber']} and "
                f"{current['windowNumber']}."
            )

    # --------------------------------------------------------
    # Window results must correspond exactly.
    # --------------------------------------------------------

    if not isinstance(
        window_results,
        list,
    ):

        errors.append(
            "window_results.json must contain a list."
        )

        return errors, normalized_windows

    result_numbers = []

    for result in window_results:

        if not isinstance(
            result,
            dict,
        ):

            errors.append(
                "A window result is not a JSON object."
            )

            continue

        number = result.get(
            "windowNumber"
        )

        if isinstance(
            number,
            int,
        ):

            result_numbers.append(
                number
            )

    expected_numbers = [
        window["windowNumber"]
        for window
        in normalized_windows
    ]

    if sorted(result_numbers) != expected_numbers:

        errors.append(
            "window_results.json window numbers do not "
            "match date_windows.json."
        )

    return errors, normalized_windows


# ============================================================
# PAGINATION VALIDATION
# ============================================================

def validate_window_pagination(
    window_result: dict[str, Any],
) -> list[str]:
    """
    Validate pagination stored inside one window result.

    IMPORTANT:
    Pagination is NOT expected in a separate pagination.json.

    The extractor stores it here:

        window_result["pagination"]
    """

    errors = []

    window_number = window_result.get(
        "windowNumber"
    )

    pagination = window_result.get(
        "pagination"
    )

    if not isinstance(
        pagination,
        list,
    ):

        return [
            f"Window {window_number}: "
            "pagination must be a list inside window_results.json."
        ]

    if not pagination:

        return [
            f"Window {window_number}: "
            "pagination list is empty."
        ]

    pages = []

    for index, page in enumerate(
        pagination,
        start=1,
    ):

        if not isinstance(
            page,
            dict,
        ):

            errors.append(
                f"Window {window_number}: "
                f"pagination entry {index} is not an object."
            )

            continue

        page_number = page.get(
            "page"
        )

        if not isinstance(
            page_number,
            int,
        ):

            errors.append(
                f"Window {window_number}: "
                f"pagination entry {index} has invalid page number."
            )

            continue

        pages.append(
            page_number
        )

        status = clean_text(
            page.get(
                "status"
            )
        ).casefold()

        if status != "success":

            errors.append(
                f"Window {window_number}, page {page_number}: "
                f"status is not 'success'."
            )

        http_status = page.get(
            "httpStatus"
        )

        if http_status != 200:

            errors.append(
                f"Window {window_number}, page {page_number}: "
                f"HTTP status is {http_status}, expected 200."
            )

        row_count = page.get(
            "rowCount"
        )

        if not isinstance(
            row_count,
            int,
        ):

            errors.append(
                f"Window {window_number}, page {page_number}: "
                "rowCount is not an integer."
            )

        elif row_count <= 0:

            errors.append(
                f"Window {window_number}, page {page_number}: "
                "rowCount must be greater than zero."
            )

        elif row_count > PAGE_SIZE:

            errors.append(
                f"Window {window_number}, page {page_number}: "
                f"rowCount={row_count} exceeds PAGE_SIZE={PAGE_SIZE}."
            )

    # --------------------------------------------------------
    # Page sequence
    # --------------------------------------------------------

    if pages:

        sorted_pages = sorted(
            pages
        )

        expected_pages = list(
            range(
                1,
                max(sorted_pages) + 1,
            )
        )

        if sorted_pages != expected_pages:

            errors.append(
                f"Window {window_number}: pagination page "
                f"sequence is incomplete. "
                f"Expected {expected_pages}, "
                f"received {sorted_pages}."
            )

        if len(
            pages
        ) != len(
            set(pages)
        ):

            errors.append(
                f"Window {window_number}: duplicate pagination "
                "page numbers detected."
            )

    # --------------------------------------------------------
    # Pagination row count reconciliation
    # --------------------------------------------------------

    valid_row_counts = [
        page.get(
            "rowCount"
        )
        for page
        in pagination
        if isinstance(
            page,
            dict,
        )
        and isinstance(
            page.get(
                "rowCount"
            ),
            int,
        )
    ]

    actual_observation_count = window_result.get(
        "observationCount"
    )

    if isinstance(
        actual_observation_count,
        int,
    ):

        pagination_row_total = sum(
            valid_row_counts
        )

        if pagination_row_total != actual_observation_count:

            errors.append(
                f"Window {window_number}: pagination row total "
                f"{pagination_row_total} does not equal "
                f"window observationCount "
                f"{actual_observation_count}."
            )

    else:

        errors.append(
            f"Window {window_number}: observationCount "
            "is not an integer."
        )

    # --------------------------------------------------------
    # Page count reconciliation
    # --------------------------------------------------------

    pages_extracted = window_result.get(
        "pagesExtracted"
    )

    if isinstance(
        pages_extracted,
        int,
    ):

        if pages_extracted != len(
            pagination
        ):

            errors.append(
                f"Window {window_number}: pagesExtracted="
                f"{pages_extracted} does not equal the "
                f"pagination list length {len(pagination)}."
            )

    else:

        errors.append(
            f"Window {window_number}: pagesExtracted "
            "is not an integer."
        )

    return errors


# ============================================================
# WINDOW RESULT VALIDATION
# ============================================================

def validate_window_results(
    window_results: Any,
    normalized_windows: list[dict[str, Any]],
) -> tuple[
    list[str],
    int,
    int,
]:
    """
    Validate all window results.

    Returns:

        errors
        total_observations
        total_pages
    """

    errors = []

    if not isinstance(
        window_results,
        list,
    ):

        return [
            "window_results.json must contain a list."
        ], 0, 0

    if len(
        window_results
    ) != len(
        normalized_windows
    ):

        errors.append(
            "window_results.json count does not match "
            "date_windows.json count."
        )

    results_by_number = {}

    for result in window_results:

        if not isinstance(
            result,
            dict,
        ):

            errors.append(
                "A window result is not a JSON object."
            )

            continue

        number = result.get(
            "windowNumber"
        )

        if not isinstance(
            number,
            int,
        ):

            errors.append(
                "A window result has invalid windowNumber."
            )

            continue

        results_by_number[
            number
        ] = result

    total_observations = 0
    total_pages = 0

    for window in normalized_windows:

        number = window[
            "windowNumber"
        ]

        result = results_by_number.get(
            number
        )

        if result is None:

            errors.append(
                f"Missing window result for window {number}."
            )

            continue

        result_start = result.get(
            "startDate"
        )

        result_end = result.get(
            "endDate"
        )

        if (
            clean_text(
                result_start
            )
            != format_date(
                window["startDate"]
            )
        ):

            errors.append(
                f"Window {number}: startDate does not match "
                "date_windows.json."
            )

        if (
            clean_text(
                result_end
            )
            != format_date(
                window["endDate"]
            )
        ):

            errors.append(
                f"Window {number}: endDate does not match "
                "date_windows.json."
            )

        historical_price_type = clean_text(
            result.get(
                "historicalPriceType"
            )
        ).upper()

        if historical_price_type != "BID":

            errors.append(
                f"Window {number}: historicalPriceType "
                "is not BID."
            )

        result_observations = result.get(
            "observations"
        )

        if not isinstance(
            result_observations,
            list,
        ):

            errors.append(
                f"Window {number}: observations is not a list."
            )

        else:

            result_observation_count = result.get(
                "observationCount"
            )

            if not isinstance(
                result_observation_count,
                int,
            ):

                errors.append(
                    f"Window {number}: observationCount "
                    "is not an integer."
                )

            elif result_observation_count != len(
                result_observations
            ):

                errors.append(
                    f"Window {number}: observationCount="
                    f"{result_observation_count} does not equal "
                    f"actual observations="
                    f"{len(result_observations)}."
                )

            total_observations += len(
                result_observations
            )

        pagination_errors = (
            validate_window_pagination(
                result
            )
        )

        errors.extend(
            pagination_errors
        )

        pagination = result.get(
            "pagination"
        )

        if isinstance(
            pagination,
            list,
        ):

            total_pages += len(
                pagination
            )

        # ----------------------------------------------------
        # Window validation object
        # ----------------------------------------------------

        validation = result.get(
            "validation"
        )

        if not isinstance(
            validation,
            dict,
        ):

            errors.append(
                f"Window {number}: validation object is missing."
            )

        else:

            if validation.get(
                "status"
            ) != "passed":

                errors.append(
                    f"Window {number}: validation status "
                    "is not passed."
                )

    return (
        errors,
        total_observations,
        total_pages,
    )


# ============================================================
# WINDOW OBSERVATION DATE VALIDATION
# ============================================================

def validate_window_observation_dates(
    window_results: list[dict[str, Any]],
    normalized_windows: list[dict[str, Any]],
) -> list[str]:

    errors = []

    windows_by_number = {
        window["windowNumber"]:
            window
        for window
        in normalized_windows
    }

    for result in window_results:

        if not isinstance(
            result,
            dict,
        ):

            continue

        number = result.get(
            "windowNumber"
        )

        window = windows_by_number.get(
            number
        )

        if window is None:

            continue

        observations = result.get(
            "observations"
        )

        if not isinstance(
            observations,
            list,
        ):

            continue

        for index, observation in enumerate(
            observations,
            start=1,
        ):

            if not isinstance(
                observation,
                dict,
            ):

                continue

            try:

                observation_date = parse_date(
                    observation.get(
                        "date"
                    )
                )

            except Exception:

                continue

            if observation_date < window["startDate"]:

                errors.append(
                    f"Window {number}, observation {index}: "
                    f"date {format_date(observation_date)} "
                    f"is before window start "
                    f"{format_date(window['startDate'])}."
                )

            if observation_date > window["endDate"]:

                errors.append(
                    f"Window {number}, observation {index}: "
                    f"date {format_date(observation_date)} "
                    f"is after window end "
                    f"{format_date(window['endDate'])}."
                )

    return errors


# ============================================================
# FULL HISTORY VALIDATION
# ============================================================

def validate_full_history(
    observations: list[dict[str, Any]],
    prudential: dict[str, Any],
    bid_history: dict[str, Any],
    summary: dict[str, Any],
) -> tuple[
    list[str],
    dict[str, Any],
]:
    """
    Validate the complete reconstructed historical BID history.
    """

    errors = []

    observation_validation = (
        validate_observations(
            observations
        )
    )

    errors.extend(
        observation_validation["errors"]
    )

    parsed_dates = (
        observation_validation["parsedDates"]
    )

    if not parsed_dates:

        return (
            errors,
            {
                "status":
                    "failed",

                "observationCount":
                    len(
                        observations
                    ),
            },
        )

    oldest_date = (
        observation_validation[
            "oldestDate"
        ]
    )

    newest_date = (
        observation_validation[
            "newestDate"
        ]
    )

    # --------------------------------------------------------
    # Prudential inception date
    # --------------------------------------------------------

    inception_raw = prudential.get(
        "inceptionDate"
    )

    try:

        inception_date = parse_date(
            inception_raw
        )

    except Exception as error:

        errors.append(
            f"Cannot validate inception date: {error}"
        )

        inception_date = None

    # --------------------------------------------------------
    # HARD INCEPTION RULE
    #
    # Oldest actual observation must be:
    #
    # inception
    # OR
    # inception + 1..7 calendar days
    #
    # Earlier than inception = FAIL.
    # Later than inception + 7 = FAIL.
    # --------------------------------------------------------

    inception_delay_days = None
    latest_allowed_oldest_date = None

    if inception_date is not None:

        latest_allowed_oldest_date = (
            inception_date
            + timedelta(days=7)
        )

        if oldest_date < inception_date:

            errors.append(
                "Historical PruAccess data contains an "
                "observation before the official Prudential "
                "inception date. "
                f"Inception={format_date(inception_date)}, "
                f"oldest={format_date(oldest_date)}."
            )

        elif oldest_date > latest_allowed_oldest_date:

            errors.append(
                "Historical PruAccess data does not reach "
                "the Prudential inception date within the "
                "permitted +7 calendar days. "
                f"Inception={format_date(inception_date)}, "
                f"oldest={format_date(oldest_date)}, "
                f"latest allowed="
                f"{format_date(latest_allowed_oldest_date)}."
            )

        inception_delay_days = (
            oldest_date
            - inception_date
        ).days

    # --------------------------------------------------------
    # Requested/current end date
    #
    # Prefer bid_history.endDate because this is the exact
    # historical extraction end date recorded by the extractor.
    # --------------------------------------------------------

    end_date_raw = bid_history.get(
        "endDate"
    )

    try:

        end_date = parse_date(
            end_date_raw
        )

    except Exception:

        end_date = None

        errors.append(
            "bid_history.json contains an invalid endDate."
        )

    if end_date is not None:

        if newest_date > end_date:

            errors.append(
                "Historical PruAccess data contains an "
                "observation after the requested historical "
                "end date. "
                f"End={format_date(end_date)}, "
                f"newest={format_date(newest_date)}."
            )

    # --------------------------------------------------------
    # Bid history observation count
    # --------------------------------------------------------

    bid_history_count = bid_history.get(
        "observationCount"
    )

    if isinstance(
        bid_history_count,
        int,
    ):

        if bid_history_count != len(
            observations
        ):

            errors.append(
                "bid_history.json observationCount does not "
                "match the actual observation count."
            )

    else:

        errors.append(
            "bid_history.json observationCount is not an integer."
        )

    # --------------------------------------------------------
    # Summary observation count
    # --------------------------------------------------------

    summary_count = summary.get(
        "historicalObservationCount"
    )

    if isinstance(
        summary_count,
        int,
    ):

        if summary_count != len(
            observations
        ):

            errors.append(
                "summary.json historicalObservationCount does "
                "not match the actual observation count."
            )

    else:

        errors.append(
            "summary.json historicalObservationCount is not "
            "an integer."
        )

    # --------------------------------------------------------
    # Summary oldest/newest observations
    # --------------------------------------------------------

    summary_oldest = summary.get(
        "oldestObservation"
    )

    summary_newest = summary.get(
        "newestObservation"
    )

    if isinstance(
        summary_oldest,
        dict,
    ):

        try:

            summary_oldest_date = parse_date(
                summary_oldest.get(
                    "date"
                )
            )

            if summary_oldest_date != oldest_date:

                errors.append(
                    "summary.json oldestObservation date does "
                    "not match reconstructed historical history."
                )

        except Exception:

            errors.append(
                "summary.json oldestObservation has an invalid date."
            )

    else:

        errors.append(
            "summary.json oldestObservation is missing."
        )

    if isinstance(
        summary_newest,
        dict,
    ):

        try:

            summary_newest_date = parse_date(
                summary_newest.get(
                    "date"
                )
            )

            if summary_newest_date != newest_date:

                errors.append(
                    "summary.json newestObservation date does "
                    "not match reconstructed historical history."
                )

        except Exception:

            errors.append(
                "summary.json newestObservation has an invalid date."
            )

    else:

        errors.append(
            "summary.json newestObservation is missing."
        )

    # --------------------------------------------------------
    # Return diagnostics
    # --------------------------------------------------------

    if (
        inception_delay_days is not None
        and inception_delay_days == 0
    ):

        inception_coverage = "exact"

    elif inception_delay_days is not None:

        inception_coverage = (
            f"inception+{inception_delay_days}-days"
        )

    else:

        inception_coverage = "unknown"

    return (
        errors,
        {
            "status":
                "passed"
                if not errors
                else "failed",

            "inceptionDate":
                (
                    format_date(inception_date)
                    if inception_date
                    else None
                ),

            "oldestHistoricalDate":
                format_date(
                    oldest_date
                ),

            "newestHistoricalDate":
                format_date(
                    newest_date
                ),

            "allowedLatestOldestDate":
                (
                    format_date(
                        latest_allowed_oldest_date
                    )
                    if latest_allowed_oldest_date
                    else None
                ),

            "maximumAllowedDaysAfterInception":
                7,

            "inceptionCoverage":
                inception_coverage,

            "inceptionDelayDays":
                inception_delay_days,

            "observationCount":
                len(
                    observations
                ),

            "duplicateDateCount":
                len(
                    observation_validation[
                        "duplicateDates"
                    ]
                ),

            "dateGapCount":
                len(
                    observation_validation[
                        "dateGaps"
                    ]
                ),

            "excessiveGapCount":
                len(
                    observation_validation[
                        "excessiveGaps"
                    ]
                ),

            "dateGaps":
                observation_validation[
                    "dateGaps"
                ],

            "excessiveGaps":
                observation_validation[
                    "excessiveGaps"
                ],
        },
    )


# ============================================================
# CURRENT BID VS LATEST PRUACCESS BID
# ============================================================

def validate_current_bid_against_latest_history(
    prudential: dict[str, Any],
    observations: list[dict[str, Any]],
) -> list[str]:
    """
    Compare the current Prudential BID with the latest
    historical PruAccess BID.

    This is a consistency check only.

    It does not replace either source.
    """

    errors = []

    if not observations:

        return errors

    current_bid_raw = prudential.get(
        "bidPrice"
    )

    if current_bid_raw in (
        None,
        "",
        "-",
    ):

        errors.append(
            "Prudential current BID is unavailable; "
            "latest PruAccess BID cannot be compared."
        )

        return errors

    try:

        current_bid = parse_numeric_price(
            current_bid_raw
        )

    except ValueError as error:

        errors.append(
            f"Prudential current BID is invalid: {error}"
        )

        return errors

    latest_observation = max(
        observations,
        key=lambda row:
            parse_date(
                row["date"]
            ),
    )

    try:

        latest_bid = parse_numeric_price(
            latest_observation[
                "bidPrice"
            ]
        )

    except ValueError as error:

        errors.append(
            f"Latest PruAccess BID is invalid: {error}"
        )

        return errors

    if not values_close(
        current_bid,
        latest_bid,
    ):

        errors.append(
            "Current Prudential BID does not match the "
            "latest PruAccess BID within the configured "
            f"tolerance of {PRICE_TOLERANCE}. "
            f"Prudential={current_bid}, "
            f"PruAccess={latest_bid}, "
            f"date={latest_observation['date']}."
        )

    return errors


# ============================================================
# FUND VALIDATION
# ============================================================

def validate_fund(
    excel_fund: dict[str, Any],
) -> dict[str, Any]:
    """
    Validate one successful extracted fund.
    """

    excel_row = excel_fund[
        "excelRow"
    ]

    errors = []
    warnings = []

    fund_directory = (
        get_fund_directory(
            excel_row
        )
    )

    if fund_directory is None:

        return {
            "excelRow":
                excel_row,

            "status":
                "failed",

            "fundDirectory":
                None,

            "errors": [
                "Individual fund output directory not found."
            ],

            "warnings":
                [],
        }

    # --------------------------------------------------------
    # Required files
    # --------------------------------------------------------

    required_file_errors = (
        validate_required_files(
            fund_directory
        )
    )

    errors.extend(
        required_file_errors
    )

    if required_file_errors:

        return {
            "excelRow":
                excel_row,

            "status":
                "failed",

            "fundDirectory":
                str(
                    fund_directory
                ),

            "errors":
                errors,

            "warnings":
                warnings,
        }

    # --------------------------------------------------------
    # Load files
    # --------------------------------------------------------

    try:

        summary = load_json(
            fund_directory
            / "summary.json"
        )

        bid_history = load_json(
            fund_directory
            / "bid_history.json"
        )

        date_windows = load_json(
            fund_directory
            / "date_windows.json"
        )

        window_diagnostics = load_json(
            fund_directory
            / "window_diagnostics.json"
        )

        window_results = load_json(
            fund_directory
            / "window_results.json"
        )

        inception_validation = load_json(
            fund_directory
            / "inception_validation.json"
        )

        prudential = load_json(
            fund_directory
            / "prudential_fund.json"
        )

    except Exception as error:

        return {
            "excelRow":
                excel_row,

            "status":
                "failed",

            "fundDirectory":
                str(
                    fund_directory
                ),

            "errors": [
                f"Unable to load fund JSON files: {error}"
            ],

            "warnings":
                warnings,
        }

    # --------------------------------------------------------
    # Structure
    # --------------------------------------------------------

    errors.extend(
        validate_summary_structure(
            summary,
            excel_fund,
        )
    )

    errors.extend(
        validate_prudential_fund(
            prudential,
            excel_fund,
        )
    )

    errors.extend(
        validate_bid_history_structure(
            bid_history
        )
    )

    # --------------------------------------------------------
    # Fund identity
    # --------------------------------------------------------

    summary_fund_name = normalize_text(
        summary.get(
            "prudentialFundName"
        )
    )

    prudential_fund_name = normalize_text(
        prudential.get(
            "fundName"
        )
    )

    if (
        summary_fund_name
        and prudential_fund_name
        and summary_fund_name
        != prudential_fund_name
    ):

        errors.append(
            "summary.json Prudential fund name does not "
            "match prudential_fund.json."
        )

    summary_identifier = normalize_text(
        summary.get(
            "fundIdentifier"
        )
    )

    prudential_identifier = normalize_text(
        prudential.get(
            "fundIdentifier"
        )
    )

    if (
        summary_identifier
        and prudential_identifier
        and summary_identifier
        != prudential_identifier
    ):

        errors.append(
            "summary.json fundIdentifier does not match "
            "prudential_fund.json."
        )

    # --------------------------------------------------------
    # Exact PruAccess name
    # --------------------------------------------------------

    excel_pruaccess_name = normalize_text(
        excel_fund.get(
            "pruAccessName"
        )
    )

    matched_pruaccess_name = normalize_text(
        summary.get(
            "matchedPruAccessName"
        )
    )

    if (
        excel_pruaccess_name
        and matched_pruaccess_name
        != excel_pruaccess_name
    ):

        errors.append(
            "Matched PruAccess fund name does not exactly "
            "match the Excel Column B reference name."
        )

    # --------------------------------------------------------
    # BID history
    # --------------------------------------------------------

    observations = []

    if isinstance(
        bid_history,
        dict,
    ):

        observations = bid_history.get(
            "observations"
        ) or []

    observation_validation = (
        validate_observations(
            observations
        )
    )

    errors.extend(
        observation_validation[
            "errors"
        ]
    )

    # --------------------------------------------------------
    # Date windows
    # --------------------------------------------------------

    window_structure_errors, normalized_windows = (
        validate_date_windows(
            date_windows,
            window_results,
        )
    )

    errors.extend(
        window_structure_errors
    )

    # --------------------------------------------------------
    # Window results + pagination
    # --------------------------------------------------------

    (
        window_result_errors,
        window_observation_total,
        window_page_total,
    ) = validate_window_results(
        window_results,
        normalized_windows,
    )

    errors.extend(
        window_result_errors
    )

    errors.extend(
        validate_window_observation_dates(
            window_results,
            normalized_windows,
        )
    )

    # --------------------------------------------------------
    # Reconciliation: window observations vs final history
    # --------------------------------------------------------

    if window_observation_total != len(
        observations
    ):

        errors.append(
            "Sum of observations across all historical "
            f"windows ({window_observation_total}) does not "
            f"equal final bid_history observation count "
            f"({len(observations)})."
        )

    # --------------------------------------------------------
    # Reconciliation: window/page totals
    # --------------------------------------------------------

    summary_pages = summary.get(
        "pagesExtracted"
    )

    if isinstance(
        summary_pages,
        int,
    ):

        if summary_pages != window_page_total:

            errors.append(
                "summary.json pagesExtracted does not match "
                "the total number of pagination pages stored "
                "inside window_results.json."
            )

    else:

        errors.append(
            "summary.json pagesExtracted is not an integer."
        )

    summary_windows = summary.get(
        "windowCount"
    )

    if isinstance(
        summary_windows,
        int,
    ):

        if summary_windows != len(
            normalized_windows
        ):

            errors.append(
                "summary.json windowCount does not match "
                "date_windows.json."
            )

    else:

        errors.append(
            "summary.json windowCount is not an integer."
        )

    summary_observation_count = summary.get(
        "historicalObservationCount"
    )

    if isinstance(
        summary_observation_count,
        int,
    ):

        if summary_observation_count != len(
            observations
        ):

            errors.append(
                "summary.json historicalObservationCount "
                "does not match bid_history.json."
            )

    # --------------------------------------------------------
    # Inception validation file
    # --------------------------------------------------------

    if not isinstance(
        inception_validation,
        dict,
    ):

        errors.append(
            "inception_validation.json must contain "
            "a JSON object."
        )

    else:

        if inception_validation.get(
            "status"
        ) != "passed":

            errors.append(
                "inception_validation.json status is not passed."
            )

        if (
            inception_validation.get(
                "maximumAllowedDaysAfterInception"
            )
            != 7
        ):

            errors.append(
                "inception_validation.json does not specify "
                "the required +7-day inception tolerance."
            )

    # --------------------------------------------------------
    # Full-history validation
    # --------------------------------------------------------

    (
        full_history_errors,
        full_history_result,
    ) = validate_full_history(
        observations,
        prudential,
        bid_history,
        summary,
    )

    errors.extend(
        full_history_errors
    )

    # --------------------------------------------------------
    # Current BID reconciliation
    # --------------------------------------------------------

    current_bid_errors = (
        validate_current_bid_against_latest_history(
            prudential,
            observations,
        )
    )

    errors.extend(
        current_bid_errors
    )

    # --------------------------------------------------------
    # Validate summary inception validation
    # --------------------------------------------------------

    summary_inception_validation = summary.get(
        "inceptionValidation"
    )

    if isinstance(
        summary_inception_validation,
        dict,
    ):

        summary_inception_status = (
            summary_inception_validation.get(
                "status"
            )
        )

        if summary_inception_status != "passed":

            errors.append(
                "summary.json inceptionValidation status "
                "is not passed."
            )

    else:

        errors.append(
            "summary.json inceptionValidation object is missing."
        )

    # --------------------------------------------------------
    # Validate extractor's window diagnostics
    # --------------------------------------------------------

    if not isinstance(
        window_diagnostics,
        list,
    ):

        errors.append(
            "window_diagnostics.json must contain a list."
        )

    else:

        if len(
            window_diagnostics
        ) != len(
            normalized_windows
        ):

            errors.append(
                "window_diagnostics.json count does not match "
                "the number of historical windows."
            )

        for diagnostic in window_diagnostics:

            if not isinstance(
                diagnostic,
                dict,
            ):

                errors.append(
                    "A window diagnostic is not a JSON object."
                )

                continue

            if diagnostic.get(
                "status"
            ) != "success":

                errors.append(
                    f"Window diagnostic "
                    f"{diagnostic.get('windowNumber')} "
                    "does not have success status."
                )

    # --------------------------------------------------------
    # Build result
    # --------------------------------------------------------

    result = {
        "excelRow":
            excel_row,

        "status":
            "passed"
            if not errors
            else "failed",

        "fundDirectory":
            str(
                fund_directory
            ),

        "fundName":
            prudential.get(
                "fundName"
            ),

        "fundIdentifier":
            prudential.get(
                "fundIdentifier"
            ),

        "fundCode":
            prudential.get(
                "fundCode"
            ),

        "observationCount":
            len(
                observations
            ),

        "windowCount":
            len(
                normalized_windows
            ),

        "pageCount":
            window_page_total,

        "oldestHistoricalDate":
            (
                format_date(
                    observation_validation[
                        "oldestDate"
                    ]
                )
                if observation_validation[
                    "oldestDate"
                ]
                else None
            ),

        "newestHistoricalDate":
            (
                format_date(
                    observation_validation[
                        "newestDate"
                    ]
                )
                if observation_validation[
                    "newestDate"
                ]
                else None
            ),

        "dateGapCount":
            len(
                observation_validation[
                    "dateGaps"
                ]
            ),

        "excessiveGapCount":
            len(
                observation_validation[
                    "excessiveGaps"
                ]
            ),

        "inceptionValidation":
            full_history_result,

        "errors":
            errors,

        "warnings":
            warnings,
    }

    return result


# ============================================================
# CONSOLIDATED ALL-FUNDS VALIDATION
# ============================================================

def validate_all_funds_file(
    all_funds: Any,
    excel_funds: list[dict[str, Any]],
) -> list[str]:

    errors = []

    if not isinstance(
        all_funds,
        dict,
    ):

        return [
            "all_funds.json must contain a JSON object."
        ]

    expected_count = len(
        excel_funds
    )

    universe_count = all_funds.get(
        "fundUniverseCount"
    )

    successful_count = all_funds.get(
        "successfulFundCount"
    )

    failed_count = all_funds.get(
        "failedFundCount"
    )

    if universe_count != expected_count:

        errors.append(
            "all_funds.json fundUniverseCount does not "
            "match the Excel master universe."
        )

    funds = all_funds.get(
        "funds"
    )

    if not isinstance(
        funds,
        list,
    ):

        errors.append(
            "all_funds.json funds must be a list."
        )

        return errors

    if successful_count != len(
        funds
    ):

        errors.append(
            "all_funds.json successfulFundCount does not "
            "match the funds list length."
        )

    failed = all_funds.get(
        "failed"
    )

    failed_funds = all_funds.get(
        "failedFunds"
    )

    if not isinstance(
        failed,
        list,
    ):

        errors.append(
            "all_funds.json failed must be a list."
        )

    if not isinstance(
        failed_funds,
        list,
    ):

        errors.append(
            "all_funds.json failedFunds must be a list."
        )

    if (
        isinstance(
            failed_count,
            int,
        )
        and isinstance(
            failed_funds,
            list,
        )
        and failed_count != len(
            failed_funds
        )
    ):

        errors.append(
            "all_funds.json failedFundCount does not "
            "match failedFunds length."
        )

    # --------------------------------------------------------
    # Excel row reconciliation
    # --------------------------------------------------------

    expected_rows = {
        fund["excelRow"]
        for fund
        in excel_funds
    }

    actual_rows = {
        fund.get(
            "excelRow"
        )
        for fund
        in funds
        if isinstance(
            fund,
            dict,
        )
    }

    missing_rows = sorted(
        expected_rows
        - actual_rows
    )

    unexpected_rows = sorted(
        actual_rows
        - expected_rows
    )

    if missing_rows:

        errors.append(
            "all_funds.json is missing successful funds "
            f"for Excel rows: {missing_rows}."
        )

    if unexpected_rows:

        errors.append(
            "all_funds.json contains funds for Excel rows "
            f"not present in the master Excel universe: "
            f"{unexpected_rows}."
        )

    return errors


# ============================================================
# CONSOLIDATED BID HISTORY VALIDATION
# ============================================================

def validate_all_bid_history_file(
    all_bid_history: Any,
    fund_results: list[dict[str, Any]],
) -> list[str]:

    errors = []

    if not isinstance(
        all_bid_history,
        dict,
    ):

        return [
            "all_bid_history.json must contain a JSON object."
        ]

    price_type = clean_text(
        all_bid_history.get(
            "priceType"
        )
    ).upper()

    if price_type != "BID":

        errors.append(
            "all_bid_history.json priceType is not BID."
        )

    funds = all_bid_history.get(
        "funds"
    )

    if not isinstance(
        funds,
        dict,
    ):

        errors.append(
            "all_bid_history.json funds must be an object."
        )

        return errors

    expected_successful_count = len(
        fund_results
    )

    actual_fund_count = all_bid_history.get(
        "fundCount"
    )

    if actual_fund_count != len(
        funds
    ):

        errors.append(
            "all_bid_history.json fundCount does not match "
            "the actual funds object count."
        )

    if actual_fund_count != expected_successful_count:

        errors.append(
            "all_bid_history.json fundCount does not match "
            "the number of successful fund results."
        )

    expected_total_observations = 0
    expected_total_windows = 0
    expected_total_pages = 0

    expected_history_keys = set()

    for result in fund_results:

        identifier = clean_text(
            result.get(
                "fundIdentifier"
            )
        )

        fund_code = clean_text(
            result.get(
                "fundCode"
            )
        )

        excel_row = result.get(
            "excelRow"
        )

        history_key = (
            identifier
            or fund_code
            or str(
                excel_row
            )
        )

        expected_history_keys.add(
            history_key
        )

        expected_total_observations += (
            result.get(
                "observationCount"
            )
            or 0
        )

        expected_total_windows += (
            result.get(
                "windowCount"
            )
            or 0
        )

        expected_total_pages += (
            result.get(
                "pageCount"
            )
            or 0
        )

        if history_key not in funds:

            errors.append(
                "all_bid_history.json is missing fund "
                f"history for key {history_key}."
            )

            continue

        history = funds[
            history_key
        ]

        if not isinstance(
            history,
            dict,
        ):

            errors.append(
                f"all_bid_history fund {history_key} "
                "is not a JSON object."
            )

            continue

        if clean_text(
            history.get(
                "priceType"
            )
        ).upper() != "BID":

            errors.append(
                f"all_bid_history fund {history_key} "
                "does not have priceType BID."
            )

        observations = history.get(
            "observations"
        )

        if not isinstance(
            observations,
            list,
        ):

            errors.append(
                f"all_bid_history fund {history_key} "
                "observations is not a list."
            )

            continue

        observation_count = history.get(
            "observationCount"
        )

        if observation_count != len(
            observations
        ):

            errors.append(
                f"all_bid_history fund {history_key} "
                "observationCount does not match observations."
            )

    actual_keys = set(
        funds.keys()
    )

    missing_keys = (
        expected_history_keys
        - actual_keys
    )

    unexpected_keys = (
        actual_keys
        - expected_history_keys
    )

    if missing_keys:

        errors.append(
            "all_bid_history.json missing expected fund "
            f"keys: {sorted(missing_keys)}."
        )

    if unexpected_keys:

        errors.append(
            "all_bid_history.json contains unexpected fund "
            f"keys: {sorted(unexpected_keys)}."
        )

    actual_total_observations = all_bid_history.get(
        "totalHistoricalBidObservations"
    )

    if actual_total_observations != expected_total_observations:

        errors.append(
            "all_bid_history.json totalHistoricalBidObservations "
            "does not reconcile with individual fund histories."
        )

    actual_total_windows = all_bid_history.get(
        "totalWindows"
    )

    if actual_total_windows != expected_total_windows:

        errors.append(
            "all_bid_history.json totalWindows does not "
            "reconcile with individual fund results."
        )

    actual_total_pages = all_bid_history.get(
        "totalPages"
    )

    if actual_total_pages != expected_total_pages:

        errors.append(
            "all_bid_history.json totalPages does not "
            "reconcile with individual fund results."
        )

    return errors


# ============================================================
# RUN SUMMARY VALIDATION
# ============================================================

def validate_run_summary(
    run_summary: Any,
    excel_funds: list[dict[str, Any]],
    fund_results: list[dict[str, Any]],
) -> list[str]:

    errors = []

    if not isinstance(
        run_summary,
        dict,
    ):

        return [
            "run_summary.json must contain a JSON object."
        ]

    expected_count = len(
        excel_funds
    )

    successful_count = len(
        fund_results
    )

    if run_summary.get(
        "fundUniverseCount"
    ) != expected_count:

        errors.append(
            "run_summary.json fundUniverseCount does not "
            "match the Excel master universe."
        )

    if run_summary.get(
        "successfulFundCount"
    ) != successful_count:

        errors.append(
            "run_summary.json successfulFundCount does not "
            "match validated successful funds."
        )

    failed_count = run_summary.get(
        "failedFundCount"
    )

    if failed_count != (
        expected_count
        - successful_count
    ):

        errors.append(
            "run_summary.json failedFundCount does not "
            "reconcile with the Excel universe and "
            "successful funds."
        )

    historical_price_type = clean_text(
        run_summary.get(
            "historicalPriceType"
        )
    ).upper()

    if historical_price_type != "BID":

        errors.append(
            "run_summary.json historicalPriceType is not BID."
        )

    if run_summary.get(
        "maximumWindowYears"
    ) != 10:

        errors.append(
            "run_summary.json maximumWindowYears is not 10."
        )

    if run_summary.get(
        "pageSize"
    ) != PAGE_SIZE:

        errors.append(
            "run_summary.json pageSize does not match "
            f"expected PAGE_SIZE={PAGE_SIZE}."
        )

    inception_rules = run_summary.get(
        "inceptionCoverageRule"
    )

    if isinstance(
        inception_rules,
        dict,
    ):

        if inception_rules.get(
            "maximumAllowedDaysAfterInception"
        ) != 7:

            errors.append(
                "run_summary.json inception rule does not "
                "specify +7 days."
            )

        if inception_rules.get(
            "dateGapsAllowed"
        ) is not True:

            errors.append(
                "run_summary.json does not record date gaps "
                "as allowed."
            )

        if inception_rules.get(
            "duplicateDatesAllowed"
        ) is not False:

            errors.append(
                "run_summary.json duplicate-date rule is incorrect."
            )

        if inception_rules.get(
            "syntheticDataAllowed"
        ) is not False:

            errors.append(
                "run_summary.json syntheticDataAllowed is not false."
            )

        if inception_rules.get(
            "interpolationAllowed"
        ) is not False:

            errors.append(
                "run_summary.json interpolationAllowed is not false."
            )

        if inception_rules.get(
            "estimationAllowed"
        ) is not False:

            errors.append(
                "run_summary.json estimationAllowed is not false."
            )

    else:

        errors.append(
            "run_summary.json inceptionCoverageRule is missing."
        )

    return errors


# ============================================================
# MAIN VALIDATION
# ============================================================

def main() -> None:

    print(
        "============================================================"
    )

    print(
        "PRUACCESS EXTRACTION VALIDATOR"
    )

    print(
        "============================================================"
    )

    print(
        "Validator version:"
    )

    print(
        " - Pagination is read from window_results.json"
    )

    print(
        " - No pagination.json is required"
    )

    print(
        " - Maximum historical BID gap: "
        f"{MAX_BID_DATE_GAP_DAYS} calendar days"
    )

    print(
        " - Maximum inception delay: +7 calendar days"
    )

    print(
        " - Synthetic/interpolated/estimated data: forbidden"
    )

    # ========================================================
    # LOAD EXCEL MASTER
    # ========================================================

    try:

        excel_funds = read_excel_funds()

    except Exception as error:

        print(
            f"\nERROR reading Excel master: {error}"
        )

        raise SystemExit(1)

    excel_universe_count = len(
        excel_funds
    )

    print(
        f"\nExcel master fund universe: "
        f"{excel_universe_count}"
    )

    # ========================================================
    # CHECK OUTPUT DIRECTORY
    # ========================================================

    if not OUTPUT_DIR.exists():

        print(
            f"\nERROR: Output directory not found: "
            f"{OUTPUT_DIR}"
        )

        raise SystemExit(1)

    if not FUNDS_OUTPUT_DIR.exists():

        print(
            f"\nERROR: Fund output directory not found: "
            f"{FUNDS_OUTPUT_DIR}"
        )

        raise SystemExit(1)

    # ========================================================
    # VALIDATE EACH FUND
    # ========================================================

    fund_results = []

    passed_count = 0
    failed_count = 0

    total_observations = 0
    total_windows = 0
    total_pages = 0

    print(
        "\n============================================================"
    )

    print(
        "INDIVIDUAL FUND VALIDATION"
    )

    print(
        "============================================================"
    )

    for index, excel_fund in enumerate(
        excel_funds,
        start=1,
    ):

        excel_row = excel_fund[
            "excelRow"
        ]

        print(
            f"\n[{index}/{excel_universe_count}] "
            f"Excel row {excel_row}"
        )

        result = validate_fund(
            excel_fund
        )

        fund_results.append(
            result
        )

        if result["status"] == "passed":

            passed_count += 1

            total_observations += (
                result.get(
                    "observationCount"
                )
                or 0
            )

            total_windows += (
                result.get(
                    "windowCount"
                )
                or 0
            )

            total_pages += (
                result.get(
                    "pageCount"
                )
                or 0
            )

            print(
                "  PASS"
            )

            print(
                f"  Fund: "
                f"{result.get('fundName')}"
            )

            print(
                f"  Identifier: "
                f"{result.get('fundIdentifier')}"
            )

            print(
                f"  Windows: "
                f"{result.get('windowCount')}"
            )

            print(
                f"  Pages: "
                f"{result.get('pageCount')}"
            )

            print(
                f"  Observations: "
                f"{result.get('observationCount')}"
            )

            print(
                f"  Oldest BID: "
                f"{result.get('oldestHistoricalDate')}"
            )

            print(
                f"  Newest BID: "
                f"{result.get('newestHistoricalDate')}"
            )

            print(
                f"  Date gaps: "
                f"{result.get('dateGapCount')}"
            )

        else:

            failed_count += 1

            print(
                "  FAILED"
            )

            for error in result.get(
                "errors",
                [],
            ):

                print(
                    f"  ERROR: {error}"
                )

    # ========================================================
    # LOAD CONSOLIDATED FILES
    # ========================================================

    consolidated_errors = []

    all_funds = None
    all_bid_history = None
    run_summary = None

    # --------------------------------------------------------
    # all_funds.json
    # --------------------------------------------------------

    if ALL_FUNDS_FILE.exists():

        try:

            all_funds = load_json(
                ALL_FUNDS_FILE
            )

            consolidated_errors.extend(
                validate_all_funds_file(
                    all_funds,
                    excel_funds,
                )
            )

        except Exception as error:

            consolidated_errors.append(
                f"Unable to validate all_funds.json: {error}"
            )

    else:

        consolidated_errors.append(
            "Missing required consolidated file: all_funds.json"
        )

    # --------------------------------------------------------
    # all_bid_history.json
    # --------------------------------------------------------

    if ALL_BID_HISTORY_FILE.exists():

        try:

            all_bid_history = load_json(
                ALL_BID_HISTORY_FILE
            )

            consolidated_errors.extend(
                validate_all_bid_history_file(
                    all_bid_history,
                    fund_results,
                )
            )

        except Exception as error:

            consolidated_errors.append(
                "Unable to validate all_bid_history.json: "
                f"{error}"
            )

    else:

        consolidated_errors.append(
            "Missing required consolidated file: "
            "all_bid_history.json"
        )

    # --------------------------------------------------------
    # run_summary.json
    # --------------------------------------------------------

    if RUN_SUMMARY_FILE.exists():

        try:

            run_summary = load_json(
                RUN_SUMMARY_FILE
            )

            consolidated_errors.extend(
                validate_run_summary(
                    run_summary,
                    excel_funds,
                    [
                        result
                        for result
                        in fund_results
                        if result["status"]
                        == "passed"
                    ],
                )
            )

        except Exception as error:

            consolidated_errors.append(
                f"Unable to validate run_summary.json: {error}"
            )

    else:

        consolidated_errors.append(
            "Missing required consolidated file: run_summary.json"
        )

    # ========================================================
    # FINAL RECONCILIATION
    # ========================================================

    if passed_count != excel_universe_count:

        consolidated_errors.append(
            "Successful validated fund count does not equal "
            "the Excel master fund universe."
        )

    if failed_count != 0:

        consolidated_errors.append(
            f"{failed_count} Excel-master fund(s) failed "
            "individual validation."
        )

    # --------------------------------------------------------
    # If extractor itself reports a failed run, fail.
    # --------------------------------------------------------

    if isinstance(
        run_summary,
        dict,
    ):

        if run_summary.get(
            "status"
        ) != "success":

            consolidated_errors.append(
                "run_summary.json reports extraction status "
                "as failed."
            )

    # --------------------------------------------------------
    # Consolidated totals
    # --------------------------------------------------------

    if isinstance(
        all_bid_history,
        dict,
    ):

        if all_bid_history.get(
            "totalHistoricalBidObservations"
        ) != total_observations:

            consolidated_errors.append(
                "Consolidated totalHistoricalBidObservations "
                "does not match validated individual funds."
            )

        if all_bid_history.get(
            "totalWindows"
        ) != total_windows:

            consolidated_errors.append(
                "Consolidated totalWindows does not match "
                "validated individual funds."
            )

        if all_bid_history.get(
            "totalPages"
        ) != total_pages:

            consolidated_errors.append(
                "Consolidated totalPages does not match "
                "validated individual funds."
            )

    # ========================================================
    # VALIDATION SUMMARY
    # ========================================================

    validation_status = (
        "passed"
        if (
            passed_count
            == excel_universe_count
            and failed_count
            == 0
            and not consolidated_errors
        )
        else "failed"
    )

    validation_summary = {
        "status":
            validation_status,

        "generatedAtUtc":
            datetime.utcnow().isoformat()
            + "Z",

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "excelFundUniverse":
            excel_universe_count,

        "successfulFunds":
            passed_count,

        "failedFunds":
            failed_count,

        "totalHistoricalBidObservations":
            total_observations,

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "rules": {
            "excelMasterUniverse":
                True,

            "hardcodedFundCount":
                False,

            "historicalPriceType":
                "BID",

            "paginationStoredInsideWindowResults":
                True,

            "paginationJsonRequired":
                False,

            "automaticTenYearWindows":
                True,

            "maximumWindowYears":
                10,

            "pageSize":
                PAGE_SIZE,

            "completePaginationValidation":
                True,

            "paginationPagesMustBeContinuous":
                True,

            "paginationPageStatusMustBeSuccess":
                True,

            "paginationRowCountsMustReconcile":
                True,

            "chronologicalReconstruction":
                True,

            "oldestObservationMustBeOnOrAfterInception":
                True,

            "maximumInceptionDelayDays":
                7,

            "firstObservationMayBeUpToSevenDaysAfterInception":
                True,

            "consecutiveHistoricalDateGapMaximumDays":
                MAX_BID_DATE_GAP_DAYS,

            "dateGapsAllowed":
                True,

            "dateGapsMustNotBeFilled":
                True,

            "duplicateDatesCauseFailure":
                True,

            "syntheticDataAllowed":
                False,

            "interpolationAllowed":
                False,

            "estimationAllowed":
                False,

            "fabricatedDataAllowed":
                False,

            "currentPrudentialBidComparedWithLatestPruAccessBid":
                True,
        },

        "fundResults":
            fund_results,

        "consolidatedErrors":
            consolidated_errors,
    }

    # ========================================================
    # SAVE VALIDATION FILE
    # ========================================================

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    VALIDATION_FILE.write_text(
        json.dumps(
            validation_summary,
            indent=2,
            ensure_ascii=False,
            default=str,
        ),
        encoding="utf-8",
    )

    # ========================================================
    # CONSOLE SUMMARY
    # ========================================================

    print(
        "\n\n"
        "============================================================"
    )

    print(
        "PRUACCESS VALIDATION COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Excel master universe: "
        f"{excel_universe_count}"
    )

    print(
        f"Funds passed: "
        f"{passed_count}"
    )

    print(
        f"Funds failed: "
        f"{failed_count}"
    )

    print(
        f"Total windows: "
        f"{total_windows}"
    )

    print(
        f"Total pages: "
        f"{total_pages}"
    )

    print(
        f"Total historical BID observations: "
        f"{total_observations}"
    )

    print(
        "\nHard rules:"
    )

    print(
        " - Historical price type: BID"
    )

    print(
        " - Pagination: validated from window_results.json"
    )

    print(
        " - pagination.json required: NO"
    )

    print(
        " - Maximum window length: 10 years"
    )

    print(
        " - Oldest observation: inception to inception +7 days"
    )

    print(
        " - Maximum consecutive BID date gap: 7 calendar days"
    )

    print(
        " - Date gaps: ALLOWED"
    )

    print(
        " - Date gaps filled: NO"
    )

    print(
        " - Synthetic data: FORBIDDEN"
    )

    print(
        " - Interpolation: FORBIDDEN"
    )

    print(
        " - Estimation: FORBIDDEN"
    )

    print(
        " - Duplicate dates: FAIL"
    )

    # ========================================================
    # FAILED FUND DETAILS
    # ========================================================

    failed_results = [
        result
        for result
        in fund_results
        if result["status"] != "passed"
    ]

    if failed_results:

        print(
            "\n============================================================"
        )

        print(
            "FAILED FUNDS"
        )

        print(
            "============================================================"
        )

        for result in failed_results:

            print(
                f"\nExcel row {result['excelRow']}:"
            )

            for error in result.get(
                "errors",
                [],
            ):

                print(
                    f" - {error}"
                )

    # ========================================================
    # CONSOLIDATED ERRORS
    # ========================================================

    if consolidated_errors:

        print(
            "\n============================================================"
        )

        print(
            "CONSOLIDATED VALIDATION ERRORS"
        )

        print(
            "============================================================"
        )

        for error in consolidated_errors:

            print(
                f" - {error}"
            )

    # ========================================================
    # FINAL GATE
    # ========================================================

    if validation_status != "passed":

        print(
            "\n============================================================"
        )

        print(
            "VALIDATION FAILED"
        )

        print(
            "============================================================"
        )

        print(
            "The extracted dataset must NOT be treated as a "
            "complete production dataset."
        )

        print(
            f"Validation file: {VALIDATION_FILE}"
        )

        raise SystemExit(1)

    print(
        "\n============================================================"
    )

    print(
        "VALIDATION PASSED"
    )

    print(
        "============================================================"
    )

    print(
        f"{passed_count} / "
        f"{excel_universe_count} Excel-master funds passed."
    )

    print(
        "All historical BID observations passed validation."
    )

    print(
        "All historical windows passed."
    )

    print(
        "All pagination pages passed."
    )

    print(
        "All pagination row counts reconciled."
    )

    print(
        "All historical date gaps are within the 7-day limit."
    )

    print(
        "All inception-coverage checks passed."
    )

    print(
        "No synthetic data was generated."
    )

    print(
        "No interpolation was used."
    )

    print(
        "No estimation was used."
    )

    print(
        "No historical date gaps were filled."
    )

    print(
        f"\nValidation file:"
        f"\n{VALIDATION_FILE}"
    )

    print(
        "\nDone."
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()
