#!/usr/bin/env python3
"""
Validate the complete all-fund PruAccess BID extraction.

Master universe:
    Funds Links.xlsm
    Column A = Prudential fund URL
    Column B = PruAccess fund name

Expected extraction output:
    output_pruaccess/
        run_summary.json
        all_funds.json
        all_bid_history.json
        funds/
            <excelRow>_<fundIdentifier>/
                summary.json
                bid_history.json
                pagination.json
                prudential_fund.json

The validator:
    - Reads the Excel universe dynamically.
    - Reads the extractor run summary.
    - Validates every successful fund.
    - Validates every individual BID history.
    - Validates pagination.
    - Validates current Prudential BID vs latest PruAccess BID.
    - Validates consolidated all_bid_history.json.
    - Fails if any of the 67 funds failed extraction.
"""

from __future__ import annotations

import json
import math
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")

RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"
ALL_FUNDS_FILE = OUTPUT_DIR / "all_funds.json"
ALL_BID_HISTORY_FILE = OUTPUT_DIR / "all_bid_history.json"

FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

VALIDATION_FILE = OUTPUT_DIR / "validation.json"

EXPECTED_FUND_COUNT = 67

PAGE_SIZE = 20

PRICE_TOLERANCE = 0.0002


# ============================================================================
# GENERAL HELPERS
# ============================================================================

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
    return clean_text(value).casefold()


def load_json(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Required file not found: {path}"
        )

    with path.open(
        "r",
        encoding="utf-8",
    ) as file:

        return json.load(file)


def save_json(
    path: Path,
    data,
):
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False,
        )

        file.write("\n")


def parse_price(value) -> float:
    text = clean_text(value)

    if not text:
        raise ValueError(
            "Empty price."
        )

    text = (
        text
        .replace(",", "")
        .replace("S$", "")
        .replace("$", "")
    )

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:
        raise ValueError(
            f"Invalid price: {value}"
        )

    price = float(
        match.group(0)
    )

    if not math.isfinite(price):
        raise ValueError(
            f"Non-finite price: {value}"
        )

    return price


def parse_date(value) -> datetime:
    text = clean_text(value)

    formats = [
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                text,
                fmt,
            )

        except ValueError:
            continue

    raise ValueError(
        f"Invalid date: {value}"
    )


def get_fund_directory(
    excel_row: int,
    fund_identifier: str,
) -> Path:

    return (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{fund_identifier}"
    )


# ============================================================================
# EXCEL MASTER UNIVERSE
# ============================================================================

def read_excel_universe() -> list[dict]:

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Required Excel file not found: "
            f"{EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    worksheet = workbook.active

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
                "excelRow": row_number,
                "prudentialUrl": prudential_url,
                "pruAccessName": pruaccess_name,
            }
        )

    workbook.close()

    return funds


# ============================================================================
# EXTRACTOR SUMMARY NORMALIZATION
# ============================================================================

def extract_successful_funds(
    run_summary: dict,
) -> list[dict]:
    """
    The extractor may store successful funds under slightly different
    structures depending on the version.

    Accept the known structures without changing the extractor itself.
    """

    candidates = [
        run_summary.get(
            "successfulFunds"
        ),
        run_summary.get(
            "funds"
        ),
        run_summary.get(
            "successful"
        ),
    ]

    for candidate in candidates:

        if isinstance(
            candidate,
            list,
        ):

            return candidate

    return []


def extract_failed_funds(
    run_summary: dict,
) -> list[dict]:
    """
    Read failed fund records from the known run-summary structures.
    """

    candidates = [
        run_summary.get(
            "failedFunds"
        ),
        run_summary.get(
            "failures"
        ),
        run_summary.get(
            "failed"
        ),
    ]

    for candidate in candidates:

        if isinstance(
            candidate,
            list,
        ):

            return candidate

    return []


def get_summary_excel_row(
    fund: dict,
):
    possible_keys = [
        "excelRow",
        "row",
        "excel_row",
    ]

    for key in possible_keys:

        value = fund.get(
            key
        )

        if value is not None:
            return value

    return None


# ============================================================================
# INDIVIDUAL FUND VALIDATION
# ============================================================================

def validate_fund(
    excel_fund: dict,
    successful_summary: dict,
) -> tuple[bool, dict]:

    excel_row = excel_fund[
        "excelRow"
    ]

    fund_name = clean_text(
        successful_summary.get(
            "fundName"
        )
    )

    fund_identifier = clean_text(
        successful_summary.get(
            "fundIdentifier"
        )
    )

    fund_code = clean_text(
        successful_summary.get(
            "fundCode"
        )
    )

    checks = []

    def add_check(
        name: str,
        passed: bool,
        details: str,
    ):

        checks.append(
            {
                "name": name,
                "passed": bool(passed),
                "details": details,
            }
        )

    # ------------------------------------------------------------------------
    # Required identity
    # ------------------------------------------------------------------------

    add_check(
        "Fund identifier present",
        bool(fund_identifier),
        (
            f"fundIdentifier='{fund_identifier}'."
        ),
    )

    if not fund_identifier:

        return False, {
            "excelRow": excel_row,
            "fundName": fund_name,
            "fundIdentifier": "",
            "fundCode": fund_code,
            "observationCount": 0,
            "checks": checks,
            "details": {},
            "passed": False,
        }

    # ------------------------------------------------------------------------
    # Fund directory
    # ------------------------------------------------------------------------

    fund_dir = get_fund_directory(
        excel_row,
        fund_identifier,
    )

    bid_history_file = (
        fund_dir / "bid_history.json"
    )

    pagination_file = (
        fund_dir / "pagination.json"
    )

    summary_file = (
        fund_dir / "summary.json"
    )

    prudential_file = (
        fund_dir / "prudential_fund.json"
    )

    required_files = [
        bid_history_file,
        pagination_file,
        summary_file,
        prudential_file,
    ]

    missing_files = [
        str(path)
        for path in required_files
        if not path.exists()
    ]

    add_check(
        "Required fund files",
        len(missing_files) == 0,
        (
            "All required files found."
            if not missing_files
            else
            "Missing: "
            + ", ".join(missing_files)
        ),
    )

    if missing_files:

        return False, {
            "excelRow": excel_row,
            "fundName": fund_name,
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "observationCount": 0,
            "checks": checks,
            "details": {
                "missingFiles": missing_files,
            },
            "passed": False,
        }

    # ------------------------------------------------------------------------
    # Load JSON files
    # ------------------------------------------------------------------------

    try:

        bid_history = load_json(
            bid_history_file
        )

        pagination = load_json(
            pagination_file
        )

        fund_summary = load_json(
            summary_file
        )

        prudential_fund = load_json(
            prudential_file
        )

        add_check(
            "Fund JSON files readable",
            True,
            "All fund JSON files parsed successfully.",
        )

    except Exception as error:

        add_check(
            "Fund JSON files readable",
            False,
            str(error),
        )

        return False, {
            "excelRow": excel_row,
            "fundName": fund_name,
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "observationCount": 0,
            "checks": checks,
            "details": {},
            "passed": False,
        }

    # ------------------------------------------------------------------------
    # Historical structure
    # ------------------------------------------------------------------------

    observations = bid_history.get(
        "observations"
    )

    if not isinstance(
        observations,
        list,
    ):

        add_check(
            "Historical observations list",
            False,
            "observations is not a list.",
        )

        return False, {
            "excelRow": excel_row,
            "fundName": fund_name,
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "observationCount": 0,
            "checks": checks,
            "details": {},
            "passed": False,
        }

    actual_count = len(
        observations
    )

    add_check(
        "Historical observations list",
        actual_count > 0,
        f"Found {actual_count} observations.",
    )

    # ------------------------------------------------------------------------
    # Source and BID type
    # ------------------------------------------------------------------------

    source = clean_text(
        bid_history.get(
            "source"
        )
    )

    price_type = clean_text(
        bid_history.get(
            "priceType"
        )
    )

    add_check(
        "Historical source",
        normalize_text(source)
        == "pruaccess",
        f"source='{source}'",
    )

    add_check(
        "Historical price type",
        normalize_text(price_type)
        == "bid",
        f"priceType='{price_type}'",
    )

    # ------------------------------------------------------------------------
    # Identity consistency
    # ------------------------------------------------------------------------

    history_name = clean_text(
        bid_history.get(
            "fundName"
        )
    )

    history_identifier = clean_text(
        bid_history.get(
            "fundIdentifier"
        )
    )

    history_code = clean_text(
        bid_history.get(
            "fundCode"
        )
    )

    add_check(
        "Fund name consistency",
        normalize_text(history_name)
        == normalize_text(fund_name),
        (
            f"Summary='{fund_name}'; "
            f"history='{history_name}'."
        ),
    )

    add_check(
        "Fund identifier consistency",
        normalize_text(history_identifier)
        == normalize_text(fund_identifier),
        (
            f"Summary='{fund_identifier}'; "
            f"history='{history_identifier}'."
        ),
    )

    add_check(
        "Fund code consistency",
        normalize_text(history_code)
        == normalize_text(fund_code),
        (
            f"Summary='{fund_code}'; "
            f"history='{history_code}'."
        ),
    )

    # ------------------------------------------------------------------------
    # Observation counts
    # ------------------------------------------------------------------------

    stored_history_count = (
        bid_history.get(
            "observationCount"
        )
    )

    stored_summary_count = (
        successful_summary.get(
            "historicalObservationCount"
        )
    )

    if stored_summary_count is None:

        stored_summary_count = (
            successful_summary.get(
                "observationCount"
            )
        )

    add_check(
        "Stored observation count",
        stored_history_count == actual_count,
        (
            f"Stored={stored_history_count}; "
            f"actual={actual_count}."
        ),
    )

    add_check(
        "Run summary observation count",
        stored_summary_count == actual_count,
        (
            f"Run summary={stored_summary_count}; "
            f"actual={actual_count}."
        ),
    )

    # ------------------------------------------------------------------------
    # Validate observations
    # ------------------------------------------------------------------------

    invalid_dates = []
    invalid_prices = []
    missing_fields = []
    parsed_observations = []

    for index, observation in enumerate(
        observations
    ):

        if not isinstance(
            observation,
            dict,
        ):

            missing_fields.append(
                {
                    "index": index,
                    "reason": (
                        "Observation is not an object."
                    ),
                }
            )

            continue

        date_value = observation.get(
            "date"
        )

        bid_value = observation.get(
            "bidPrice"
        )

        if (
            date_value is None
            or bid_value is None
            or clean_text(date_value) == ""
            or clean_text(bid_value) == ""
        ):

            missing_fields.append(
                {
                    "index": index,
                    "reason": (
                        "Missing date or bidPrice."
                    ),
                    "observation": observation,
                }
            )

            continue

        parsed_date = None
        parsed_price = None

        try:

            parsed_date = parse_date(
                date_value
            )

        except ValueError as error:

            invalid_dates.append(
                {
                    "index": index,
                    "date": date_value,
                    "error": str(error),
                }
            )

        try:

            parsed_price = parse_price(
                bid_value
            )

            if parsed_price <= 0:

                invalid_prices.append(
                    {
                        "index": index,
                        "bidPrice": bid_value,
                        "error": (
                            "Price must be greater than zero."
                        ),
                    }
                )

        except ValueError as error:

            invalid_prices.append(
                {
                    "index": index,
                    "bidPrice": bid_value,
                    "error": str(error),
                }
            )

        if (
            parsed_date is not None
            and parsed_price is not None
        ):

            parsed_observations.append(
                {
                    "index": index,
                    "date": parsed_date,
                    "dateRaw": clean_text(
                        date_value
                    ),
                    "price": parsed_price,
                    "priceRaw": clean_text(
                        bid_value
                    ),
                }
            )

    add_check(
        "All dates valid",
        len(invalid_dates) == 0,
        f"Invalid dates={len(invalid_dates)}.",
    )

    add_check(
        "All BID prices numeric",
        len(invalid_prices) == 0,
        f"Invalid prices={len(invalid_prices)}.",
    )

    add_check(
        "No missing observation fields",
        len(missing_fields) == 0,
        (
            f"Missing/invalid records="
            f"{len(missing_fields)}."
        ),
    )

    # ------------------------------------------------------------------------
    # Date ordering
    # ------------------------------------------------------------------------

    ordering_violations = []

    for index in range(
        1,
        len(parsed_observations),
    ):

        previous = parsed_observations[
            index - 1
        ]

        current = parsed_observations[
            index
        ]

        if current["date"] >= previous["date"]:

            ordering_violations.append(
                {
                    "index": current["index"],
                    "previousDate": (
                        previous["dateRaw"]
                    ),
                    "currentDate": (
                        current["dateRaw"]
                    ),
                }
            )

    add_check(
        "Dates strictly descending",
        len(ordering_violations) == 0,
        (
            f"Ordering violations="
            f"{len(ordering_violations)}."
        ),
    )

    # ------------------------------------------------------------------------
    # Duplicate dates
    # ------------------------------------------------------------------------

    date_counts = {}

    for observation in observations:

        date_value = clean_text(
            observation.get(
                "date"
            )
        )

        if not date_value:
            continue

        date_counts[date_value] = (
            date_counts.get(
                date_value,
                0,
            )
            + 1
        )

    duplicate_dates = {
        date: count
        for date, count in date_counts.items()
        if count > 1
    }

    add_check(
        "No duplicate dates",
        len(duplicate_dates) == 0,
        (
            f"Duplicate dates="
            f"{len(duplicate_dates)}."
        ),
    )

    # ------------------------------------------------------------------------
    # Duplicate records
    # ------------------------------------------------------------------------

    record_counts = {}

    for observation in observations:

        date_value = clean_text(
            observation.get(
                "date"
            )
        )

        price_value = clean_text(
            observation.get(
                "bidPrice"
            )
        )

        key = (
            date_value,
            price_value,
        )

        record_counts[key] = (
            record_counts.get(
                key,
                0,
            )
            + 1
        )

    duplicate_records = {
        f"{key[0]}|{key[1]}": count
        for key, count in record_counts.items()
        if count > 1
    }

    add_check(
        "No duplicate date/price records",
        len(duplicate_records) == 0,
        (
            f"Duplicate records="
            f"{len(duplicate_records)}."
        ),
    )

    # ------------------------------------------------------------------------
    # Newest / oldest
    # ------------------------------------------------------------------------

    newest_observation = (
        observations[0]
        if observations
        else None
    )

    oldest_observation = (
        observations[-1]
        if observations
        else None
    )

    add_check(
        "Newest historical observation identified",
        newest_observation is not None,
        (
            (
                f"{newest_observation.get('date')} "
                f"BID="
                f"{newest_observation.get('bidPrice')}."
            )
            if newest_observation
            else
            "No observations."
        ),
    )

    add_check(
        "Oldest historical observation identified",
        oldest_observation is not None,
        (
            (
                f"{oldest_observation.get('date')} "
                f"BID="
                f"{oldest_observation.get('bidPrice')}."
            )
            if oldest_observation
            else
            "No observations."
        ),
    )

    # ------------------------------------------------------------------------
    # Inception validation
    # ------------------------------------------------------------------------

    inception_value = clean_text(
        successful_summary.get(
            "inceptionDate"
        )
    )

    if not inception_value:

        inception_value = clean_text(
            fund_summary.get(
                "inceptionDate"
            )
        )

    inception_date = None

    if inception_value:

        try:

            inception_date = parse_date(
                inception_value
            )

        except ValueError:
            inception_date = None

    if (
        inception_date is not None
        and oldest_observation is not None
    ):

        try:

            oldest_date = parse_date(
                oldest_observation.get(
                    "date"
                )
            )

            reaches_inception = (
                oldest_date.date()
                <= inception_date.date()
            )

        except Exception:

            reaches_inception = False

        add_check(
            "Historical data reaches inception",
            reaches_inception,
            (
                f"Inception={inception_value}; "
                f"oldest observation="
                f"{oldest_observation.get('date')}."
            ),
        )

    else:

        add_check(
            "Historical data reaches inception",
            False,
            (
                "Could not determine inception date."
            ),
        )

    # ------------------------------------------------------------------------
    # Suspicious generated-data fields
    # ------------------------------------------------------------------------

    suspicious_fields = []

    suspicious_keywords = [
        "synthetic",
        "interpolat",
        "simulat",
        "generated",
        "estimated",
        "illustrative",
        "mock",
        "placeholder",
        "carryforward",
        "carry-forward",
        "forecast",
    ]

    def inspect_object(
        obj,
        location: str,
    ):

        if isinstance(
            obj,
            dict,
        ):

            for key, value in obj.items():

                key_text = normalize_text(
                    key
                )

                if any(
                    keyword in key_text
                    for keyword in suspicious_keywords
                ):

                    suspicious_fields.append(
                        {
                            "location": location,
                            "field": key,
                            "value": value,
                        }
                    )

                inspect_object(
                    value,
                    f"{location}.{key}",
                )

        elif isinstance(
            obj,
            list,
        ):

            for index, value in enumerate(
                obj
            ):

                inspect_object(
                    value,
                    f"{location}[{index}]",
                )

    inspect_object(
        bid_history,
        "bid_history",
    )

    add_check(
        "No synthetic/interpolated fields",
        len(suspicious_fields) == 0,
        (
            "No generated-data fields detected."
            if not suspicious_fields
            else
            f"Suspicious fields="
            f"{len(suspicious_fields)}."
        ),
    )

    # ------------------------------------------------------------------------
    # Pagination
    # ------------------------------------------------------------------------

    pages = pagination.get(
        "pages"
    )

    if not isinstance(
        pages,
        list,
    ):

        pages = []

    stored_page_count = pagination.get(
        "pagesExtracted"
    )

    actual_page_count = len(
        pages
    )

    add_check(
        "Pagination JSON structure",
        actual_page_count > 0,
        (
            f"Pages found="
            f"{actual_page_count}."
        ),
    )

    add_check(
        "Pagination page count",
        stored_page_count == actual_page_count,
        (
            f"Stored={stored_page_count}; "
            f"actual={actual_page_count}."
        ),
    )

    pagination_row_total = 0

    page_number_violations = []

    full_page_violations = []

    final_page_valid = False

    for index, page in enumerate(
        pages
    ):

        if not isinstance(
            page,
            dict,
        ):
            continue

        expected_page_number = (
            index + 1
        )

        actual_page_number = page.get(
            "page"
        )

        if actual_page_number != expected_page_number:

            page_number_violations.append(
                {
                    "expected": expected_page_number,
                    "actual": actual_page_number,
                }
            )

        row_count = page.get(
            "rowCount"
        )

        if not isinstance(
            row_count,
            int,
        ):
            continue

        pagination_row_total += row_count

        is_final_page = (
            index
            == len(pages) - 1
        )

        if not is_final_page:

            if row_count != PAGE_SIZE:

                full_page_violations.append(
                    {
                        "page": actual_page_number,
                        "rowCount": row_count,
                    }
                )

        else:

            final_page_valid = (
                row_count > 0
                and row_count <= PAGE_SIZE
            )

    add_check(
        "Pagination page numbering",
        len(page_number_violations) == 0,
        (
            "Pages are sequential."
            if not page_number_violations
            else
            f"Violations="
            f"{len(page_number_violations)}."
        ),
    )

    add_check(
        "Pagination full pages contain 20 rows",
        len(full_page_violations) == 0,
        (
            f"Full pages checked="
            f"{max(actual_page_count - 1, 0)}."
            if not full_page_violations
            else
            f"Violations="
            f"{len(full_page_violations)}."
        ),
    )

    add_check(
        "Final pagination page is valid",
        final_page_valid,
        (
            (
                f"Final page rows="
                f"{pages[-1].get('rowCount')}."
            )
            if pages
            else
            "No pages."
        ),
    )

    add_check(
        "Pagination row total matches observations",
        pagination_row_total == actual_count,
        (
            f"Pagination rows="
            f"{pagination_row_total}; "
            f"observations="
            f"{actual_count}."
        ),
    )

    # ------------------------------------------------------------------------
    # Current Prudential BID vs latest PruAccess BID
    # ------------------------------------------------------------------------

    current_prudential_bid = None

    if isinstance(
        prudential_fund,
        dict,
    ):

        # Most likely structure.
        if prudential_fund.get(
            "bidPrice"
        ) is not None:

            current_prudential_bid = (
                prudential_fund.get(
                    "bidPrice"
                )
            )

        else:

            data = prudential_fund.get(
                "data"
            )

            if isinstance(
                data,
                dict,
            ):

                current_prudential_bid = (
                    data.get(
                        "bidPrice"
                    )
                )

    latest_pruaccess_bid = None
    latest_pruaccess_date = None

    if newest_observation:

        latest_pruaccess_bid = (
            newest_observation.get(
                "bidPrice"
            )
        )

        latest_pruaccess_date = (
            newest_observation.get(
                "date"
            )
        )

    bid_difference = None
    bid_match_passed = False

    try:

        if (
            current_prudential_bid
            is not None
            and latest_pruaccess_bid
            is not None
        ):

            prudential_price = parse_price(
                current_prudential_bid
            )

            pruaccess_price = parse_price(
                latest_pruaccess_bid
            )

            bid_difference = abs(
                prudential_price
                - pruaccess_price
            )

            bid_match_passed = (
                bid_difference
                <= PRICE_TOLERANCE
            )

    except Exception:

        bid_match_passed = False

    add_check(
        "Current Prudential BID matches latest PruAccess BID",
        bid_match_passed,
        (
            f"Prudential current BID="
            f"{current_prudential_bid}; "
            f"PruAccess latest BID="
            f"{latest_pruaccess_bid}; "
            f"difference="
            f"{bid_difference}; "
            f"tolerance="
            f"{PRICE_TOLERANCE}."
        ),
    )

    # ------------------------------------------------------------------------
    # Overall fund result
    # ------------------------------------------------------------------------

    passed = all(
        check["passed"]
        for check in checks
    )

    result = {
        "excelRow": excel_row,
        "fundName": fund_name,
        "fundIdentifier": fund_identifier,
        "fundCode": fund_code,
        "observationCount": actual_count,
        "newestObservation": newest_observation,
        "oldestObservation": oldest_observation,
        "checks": checks,
        "details": {
            "invalidDates": invalid_dates,
            "invalidPrices": invalid_prices,
            "missingFields": missing_fields,
            "orderingViolations": ordering_violations,
            "duplicateDates": duplicate_dates,
            "duplicateRecords": duplicate_records,
            "suspiciousGeneratedFields": suspicious_fields,
            "paginationRowTotal": pagination_row_total,
            "paginationPageCount": actual_page_count,
            "currentPrudentialBid": current_prudential_bid,
            "latestPruAccessBid": latest_pruaccess_bid,
            "latestPruAccessDate": latest_pruaccess_date,
            "bidDifference": bid_difference,
        },
        "passed": passed,
    }

    return passed, result


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:

    print("=" * 70)
    print(
        "PruAccess All-Fund Historical BID Price Validation"
    )
    print("=" * 70)

    try:

        # ====================================================================
        # Load master Excel universe
        # ====================================================================

        excel_funds = read_excel_universe()

        expected_count = len(
            excel_funds
        )

        print()
        print(
            f"Excel fund universe: "
            f"{expected_count}"
        )

        # ====================================================================
        # Load extractor files
        # ====================================================================

        run_summary = load_json(
            RUN_SUMMARY_FILE
        )

        all_funds = load_json(
            ALL_FUNDS_FILE
        )

        all_bid_history = load_json(
            ALL_BID_HISTORY_FILE
        )

        # ====================================================================
        # Read successful / failed results
        # ====================================================================

        successful_funds = (
            extract_successful_funds(
                run_summary
            )
        )

        failed_fund_records = (
            extract_failed_funds(
                run_summary
            )
        )

        successful_count = len(
            successful_funds
        )

        failed_count = len(
            failed_fund_records
        )

        # If explicit counts exist, display them for comparison.
        reported_successful_count = (
            run_summary.get(
                "successfulFundCount"
            )
        )

        reported_failed_count = (
            run_summary.get(
                "failedFundCount"
            )
        )

        print(
            f"Run successful funds: "
            f"{successful_count}"
        )

        print(
            f"Run failed funds: "
            f"{failed_count}"
        )

        if (
            reported_successful_count
            != successful_count
        ):

            print(
                "WARNING: run_summary successfulFundCount "
                f"is {reported_successful_count}, "
                f"but successful fund records found="
                f"{successful_count}"
            )

        if (
            reported_failed_count
            != failed_count
        ):

            print(
                "WARNING: run_summary failedFundCount "
                f"is {reported_failed_count}, "
                f"but failed fund records found="
                f"{failed_count}"
            )

        # ====================================================================
        # Top-level checks
        # ====================================================================

        top_level_checks = []

        def top_check(
            name: str,
            passed: bool,
            details: str,
        ):

            top_level_checks.append(
                {
                    "name": name,
                    "passed": bool(passed),
                    "details": details,
                }
            )

        top_check(
            "Excel fund universe count",
            expected_count == EXPECTED_FUND_COUNT,
            (
                f"Found {expected_count} funds "
                f"in Funds Links.xlsm; "
                f"expected {EXPECTED_FUND_COUNT}."
            ),
        )

        top_check(
            "Run summary universe count",
            run_summary.get(
                "fundUniverseCount"
            ) == expected_count,
            (
                f"Run summary="
                f"{run_summary.get('fundUniverseCount')}; "
                f"Excel="
                f"{expected_count}."
            ),
        )

        top_check(
            "All funds extracted successfully",
            successful_count == expected_count,
            (
                f"Successful records="
                f"{successful_count}; "
                f"expected="
                f"{expected_count}."
            ),
        )

        top_check(
            "No failed funds",
            failed_count == 0,
            (
                f"Failed records="
                f"{failed_count}."
            ),
        )

        # ====================================================================
        # Build successful lookup by Excel row
        # ====================================================================

        successful_by_row = {}

        duplicate_success_rows = []

        for fund in successful_funds:

            row = get_summary_excel_row(
                fund
            )

            if row is None:
                continue

            if row in successful_by_row:

                duplicate_success_rows.append(
                    row
                )

            successful_by_row[row] = fund

        missing_excel_rows = [
            fund["excelRow"]
            for fund in excel_funds
            if fund["excelRow"]
            not in successful_by_row
        ]

        top_check(
            "Every Excel fund has successful result",
            len(missing_excel_rows) == 0,
            (
                "All Excel rows have successful results."
                if not missing_excel_rows
                else
                "Missing Excel rows: "
                + ", ".join(
                    str(row)
                    for row in missing_excel_rows
                )
            ),
        )

        top_check(
            "No duplicate successful Excel rows",
            len(duplicate_success_rows) == 0,
            (
                "No duplicate successful rows."
                if not duplicate_success_rows
                else
                f"Duplicate rows="
                f"{duplicate_success_rows}."
            ),
        )

        # ====================================================================
        # Validate individual funds
        # ====================================================================

        print()
        print(
            "Validating individual funds..."
        )

        fund_results = []

        passed_funds = 0
        failed_validation_funds = 0

        for index, excel_fund in enumerate(
            excel_funds,
            start=1,
        ):

            excel_row = excel_fund[
                "excelRow"
            ]

            print(
                f"  [{index}/{expected_count}] "
                f"Excel row {excel_row}: ",
                end="",
                flush=True,
            )

            successful_summary = (
                successful_by_row.get(
                    excel_row
                )
            )

            if successful_summary is None:

                result = {
                    "excelRow": excel_row,
                    "fundName": "",
                    "fundIdentifier": "",
                    "fundCode": "",
                    "observationCount": 0,
                    "checks": [
                        {
                            "name": (
                                "Successful extraction exists"
                            ),
                            "passed": False,
                            "details": (
                                "No successful extraction "
                                "record exists for this "
                                "Excel row."
                            ),
                        }
                    ],
                    "details": {},
                    "passed": False,
                }

                fund_results.append(
                    result
                )

                failed_validation_funds += 1

                print(
                    "FAILED"
                )

                continue

            passed, result = validate_fund(
                excel_fund,
                successful_summary,
            )

            fund_results.append(
                result
            )

            if passed:

                passed_funds += 1

                print(
                    f"PASS "
                    f"({result['observationCount']} observations)"
                )

            else:

                failed_validation_funds += 1

                print(
                    "FAILED"
                )

        # ====================================================================
        # Consolidated history
        # ====================================================================

        print()
        print(
            "Validating consolidated BID history..."
        )

        consolidated_funds = (
            all_bid_history.get(
                "funds"
            )
        )

        if not isinstance(
            consolidated_funds,
            dict,
        ):

            consolidated_funds = {}

        consolidated_mismatches = []

        consolidated_total = 0

        for result in fund_results:

            if not result.get(
                "fundIdentifier"
            ):

                continue

            identifier = result[
                "fundIdentifier"
            ]

            consolidated = (
                consolidated_funds.get(
                    identifier
                )
            )

            if consolidated is None:

                consolidated_mismatches.append(
                    {
                        "fundIdentifier": identifier,
                        "reason": (
                            "Missing from consolidated "
                            "all_bid_history.json."
                        ),
                    }
                )

                continue

            consolidated_observations = (
                consolidated.get(
                    "observations"
                )
            )

            if not isinstance(
                consolidated_observations,
                list,
            ):

                consolidated_mismatches.append(
                    {
                        "fundIdentifier": identifier,
                        "reason": (
                            "Consolidated observations "
                            "is not a list."
                        ),
                    }
                )

                continue

            consolidated_count = len(
                consolidated_observations
            )

            consolidated_total += (
                consolidated_count
            )

            if consolidated_count != result[
                "observationCount"
            ]:

                consolidated_mismatches.append(
                    {
                        "fundIdentifier": identifier,
                        "reason": (
                            "Observation count mismatch."
                        ),
                        "consolidated": (
                            consolidated_count
                        ),
                        "individual": (
                            result[
                                "observationCount"
                            ]
                        ),
                    }
                )

        top_check(
            "Consolidated BID history fund count",
            len(consolidated_funds)
            == successful_count,
            (
                f"Consolidated funds="
                f"{len(consolidated_funds)}; "
                f"successful funds="
                f"{successful_count}."
            ),
        )

        top_check(
            "Consolidated BID history matches individual files",
            len(consolidated_mismatches) == 0,
            (
                "All consolidated histories match."
                if not consolidated_mismatches
                else
                f"Mismatches="
                f"{len(consolidated_mismatches)}."
            ),
        )

        reported_total = (
            run_summary.get(
                "totalHistoricalBidObservations"
            )
        )

        top_check(
            "Consolidated observation total",
            consolidated_total == reported_total,
            (
                f"Consolidated="
                f"{consolidated_total}; "
                f"run summary="
                f"{reported_total}."
            ),
        )

        # ====================================================================
        # all_funds.json
        # ====================================================================

        all_funds_list = all_funds.get(
            "funds"
        )

        if not isinstance(
            all_funds_list,
            list,
        ):

            all_funds_list = []

        top_check(
            "all_funds.json fund count",
            len(all_funds_list)
            == successful_count,
            (
                f"all_funds.json="
                f"{len(all_funds_list)}; "
                f"successful="
                f"{successful_count}."
            ),
        )

        # ====================================================================
        # Observation total
        # ====================================================================

        individual_total = sum(
            result.get(
                "observationCount",
                0,
            )
            for result in fund_results
        )

        top_check(
            "Individual observation total",
            individual_total == reported_total,
            (
                f"Individual="
                f"{individual_total}; "
                f"run summary="
                f"{reported_total}."
            ),
        )

        # ====================================================================
        # Overall status
        # ====================================================================

        top_level_passed = all(
            check["passed"]
            for check in top_level_checks
        )

        overall_passed = (
            top_level_passed
            and passed_funds == expected_count
            and failed_validation_funds == 0
        )

        status = (
            "PASS"
            if overall_passed
            else
            "FAIL"
        )

        # ====================================================================
        # Write validation.json
        # ====================================================================

        validation = {
            "status": status,

            "validatedAtUtc": (
                datetime.now(
                    timezone.utc
                ).isoformat()
            ),

            "excelFile": str(
                EXCEL_FILE
            ),

            "fundUniverseCount": expected_count,

            "successfulExtractionCount": (
                successful_count
            ),

            "failedExtractionCount": (
                failed_count
            ),

            "validatedPassCount": (
                passed_funds
            ),

            "validatedFailCount": (
                failed_validation_funds
            ),

            "totalHistoricalBidObservations": (
                individual_total
            ),

            "reportedHistoricalBidObservations": (
                reported_total
            ),

            "topLevelChecks": top_level_checks,

            "fundResults": fund_results,

            "failedExtractionRecords": (
                failed_fund_records
            ),

            "details": {
                "missingExcelRows": (
                    missing_excel_rows
                ),
                "duplicateSuccessfulRows": (
                    duplicate_success_rows
                ),
                "consolidatedMismatches": (
                    consolidated_mismatches
                ),
            },
        }

        save_json(
            VALIDATION_FILE,
            validation,
        )

        # ====================================================================
        # Final console report
        # ====================================================================

        print()
        print("=" * 70)
        print(
            f"VALIDATION {status}"
        )
        print("=" * 70)

        print(
            f"Excel funds: "
            f"{expected_count}"
        )

        print(
            f"Successful extraction records: "
            f"{successful_count}"
        )

        print(
            f"Failed extraction records: "
            f"{failed_count}"
        )

        print(
            f"Funds passed validation: "
            f"{passed_funds}"
        )

        print(
            f"Funds failed validation: "
            f"{failed_validation_funds}"
        )

        print(
            f"Historical BID observations: "
            f"{individual_total}"
        )

        print()
        print(
            "Top-level checks:"
        )

        for check in top_level_checks:

            marker = (
                "PASS"
                if check["passed"]
                else
                "FAIL"
            )

            print(
                f"  [{marker}] "
                f"{check['name']}: "
                f"{check['details']}"
            )

        # ====================================================================
        # Failed funds
        # ====================================================================

        if failed_validation_funds:

            print()
            print(
                "FAILED FUNDS:"
            )

            for result in fund_results:

                if result.get(
                    "passed"
                ):
                    continue

                print(
                    f"  Excel row "
                    f"{result.get('excelRow')}: "
                    f"{result.get('fundName')}"
                )

                for check in result.get(
                    "checks",
                    [],
                ):

                    if not check.get(
                        "passed"
                    ):

                        print(
                            f"    - "
                            f"{check.get('name')}: "
                            f"{check.get('details')}"
                        )

        print()
        print(
            f"Validation file: "
            f"{VALIDATION_FILE}"
        )

        print(
            "=" * 70
        )

        if overall_passed:

            print(
                "ALL 67 FUNDS PASSED VALIDATION."
            )

            return 0

        print(
            "VALIDATION FAILED."
        )

        return 1

    except Exception as error:

        print()
        print("=" * 70)
        print(
            "VALIDATION ERROR"
        )
        print("=" * 70)

        print(
            f"{type(error).__name__}: "
            f"{error}"
        )

        return 1


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":

    sys.exit(
        main()
    )
