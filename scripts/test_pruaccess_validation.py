#!/usr/bin/env python3
"""
Validate the complete all-fund PruAccess BID extraction.

Validation rules
----------------
1. Funds Links.xlsm is the master fund universe.
2. Every Excel fund must have a successful extraction.
3. Historical prices must come from PruAccess.
4. Historical price type must be BID.
5. Every observation must contain a valid date and BID price.
6. Dates must be strictly descending.
7. No duplicate dates.
8. No duplicate date/price records.
9. Pagination counts must match observations.
10. No synthetic/interpolated/generated data fields.
11. Latest PruAccess BID must be reasonably consistent with the
    current Prudential BID.
12. Consolidated all_bid_history.json must agree with individual
    fund bid_history.json files.
13. The validator exits with code 1 if any required check fails.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCEL_FILE = Path("Funds Links.xlsm")
OUTPUT_DIR = Path("output_pruaccess")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"
ALL_FUNDS_FILE = OUTPUT_DIR / "all_funds.json"
ALL_BID_HISTORY_FILE = OUTPUT_DIR / "all_bid_history.json"

VALIDATION_FILE = OUTPUT_DIR / "validation.json"

PAGE_SIZE = 20

PRICE_TOLERANCE = 0.0002


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_json(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Required file not found: {path}"
        )

    with path.open(
        "r",
        encoding="utf-8",
    ) as handle:
        return json.load(handle)


def clean_text(value) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_text(value) -> str:
    return clean_text(value).casefold()


def parse_price(value) -> float:
    text = clean_text(value)

    text = text.replace(",", "")
    text = text.replace("$", "")
    text = text.replace("S$", "")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:
        raise ValueError(
            f"Invalid price: {value}"
        )

    return float(match.group(0))


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


def read_excel_universe() -> list[dict]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Required Excel file not found: {EXCEL_FILE}"
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


def get_fund_directory(
    excel_row: int,
    fund_identifier: str,
) -> Path:
    return (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{fund_identifier}"
    )


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_fund(
    excel_fund: dict,
    successful_summary: dict,
) -> tuple[bool, dict]:
    """
    Validate one successfully extracted fund.
    """

    excel_row = excel_fund["excelRow"]

    fund_name = clean_text(
        successful_summary.get("fundName")
    )

    fund_identifier = clean_text(
        successful_summary.get("fundIdentifier")
    )

    fund_code = clean_text(
        successful_summary.get("fundCode")
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

    # -----------------------------------------------------------------------
    # Basic identity
    # -----------------------------------------------------------------------

    if not fund_identifier:
        add_check(
            "Fund identifier present",
            False,
            "fundIdentifier is empty.",
        )

        return False, {
            "excelRow": excel_row,
            "fundName": fund_name,
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "checks": checks,
        }

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

    # -----------------------------------------------------------------------
    # Required files
    # -----------------------------------------------------------------------

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
        not missing_files,
        (
            "All required fund files found."
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
            "checks": checks,
        }

    # -----------------------------------------------------------------------
    # Load fund files
    # -----------------------------------------------------------------------

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
            "checks": checks,
        }

    add_check(
        "Fund JSON files readable",
        True,
        "All fund JSON files parsed successfully.",
    )

    # -----------------------------------------------------------------------
    # BID history structure
    # -----------------------------------------------------------------------

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
            "checks": checks,
        }

    add_check(
        "Historical observations list",
        True,
        f"Found {len(observations)} observations.",
    )

    # -----------------------------------------------------------------------
    # Source
    # -----------------------------------------------------------------------

    source = clean_text(
        bid_history.get("source")
    )

    add_check(
        "Historical source",
        normalize_text(source) == "pruaccess",
        f"source='{source}'",
    )

    price_type = clean_text(
        bid_history.get("priceType")
    )

    add_check(
        "Historical price type",
        normalize_text(price_type) == "bid",
        f"priceType='{price_type}'",
    )

    # -----------------------------------------------------------------------
    # Identity consistency
    # -----------------------------------------------------------------------

    history_fund_name = clean_text(
        bid_history.get("fundName")
    )

    history_identifier = clean_text(
        bid_history.get("fundIdentifier")
    )

    history_fund_code = clean_text(
        bid_history.get("fundCode")
    )

    add_check(
        "Fund name consistency",
        normalize_text(history_fund_name)
        == normalize_text(fund_name),
        (
            f"Summary='{fund_name}'; "
            f"history='{history_fund_name}'"
        ),
    )

    add_check(
        "Fund identifier consistency",
        normalize_text(history_identifier)
        == normalize_text(fund_identifier),
        (
            f"Summary='{fund_identifier}'; "
            f"history='{history_identifier}'"
        ),
    )

    add_check(
        "Fund code consistency",
        normalize_text(history_fund_code)
        == normalize_text(fund_code),
        (
            f"Summary='{fund_code}'; "
            f"history='{history_fund_code}'"
        ),
    )

    # -----------------------------------------------------------------------
    # Observation count
    # -----------------------------------------------------------------------

    actual_count = len(observations)

    stored_count = bid_history.get(
        "observationCount"
    )

    add_check(
        "Stored observation count",
        stored_count == actual_count,
        (
            f"Stored={stored_count}; "
            f"actual={actual_count}."
        ),
    )

    summary_count = successful_summary.get(
        "observationCount"
    )

    add_check(
        "Run summary observation count",
        summary_count == actual_count,
        (
            f"Run summary={summary_count}; "
            f"actual={actual_count}."
        ),
    )

    # -----------------------------------------------------------------------
    # Validate observations
    # -----------------------------------------------------------------------

    invalid_dates = []
    invalid_prices = []
    missing_fields = []
    parsed_dates = []
    duplicate_dates = {}
    duplicate_records = {}

    previous_date = None

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
                    "reason": "Observation is not an object.",
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
                    "reason": "Missing date or bidPrice.",
                    "observation": observation,
                }
            )

            continue

        # Date validation.
        try:

            parsed_date = parse_date(
                date_value
            )

            parsed_dates.append(
                parsed_date
            )

        except ValueError as error:

            invalid_dates.append(
                {
                    "index": index,
                    "date": date_value,
                    "error": str(error),
                }
            )

            continue

        # Price validation.
        try:

            parse_price(
                bid_value
            )

        except ValueError as error:

            invalid_prices.append(
                {
                    "index": index,
                    "bidPrice": bid_value,
                    "error": str(error),
                }
            )

        # Ordering validation.
        if previous_date is not None:

            if parsed_date >= previous_date:

                # Ordering violation is recorded later
                # using the parsed date list.

                pass

        previous_date = parsed_date

    # -----------------------------------------------------------------------
    # Date ordering
    # -----------------------------------------------------------------------

    ordering_violations = []

    for index in range(
        1,
        len(parsed_dates),
    ):

        if parsed_dates[index] >= parsed_dates[index - 1]:

            ordering_violations.append(
                {
                    "index": index,
                    "previousDate": (
                        parsed_dates[index - 1]
                        .strftime("%d-%b-%Y")
                    ),
                    "currentDate": (
                        parsed_dates[index]
                        .strftime("%d-%b-%Y")
                    ),
                }
            )

    # -----------------------------------------------------------------------
    # Duplicate dates
    # -----------------------------------------------------------------------

    date_counts = {}

    for observation in observations:

        date_value = clean_text(
            observation.get("date")
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

    # -----------------------------------------------------------------------
    # Duplicate date/price records
    # -----------------------------------------------------------------------

    record_counts = {}

    for observation in observations:

        date_value = clean_text(
            observation.get("date")
        )

        price_value = clean_text(
            observation.get("bidPrice")
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

    add_check(
        "Dates strictly descending",
        len(ordering_violations) == 0,
        (
            f"Ordering violations="
            f"{len(ordering_violations)}."
        ),
    )

    add_check(
        "No duplicate dates",
        len(duplicate_dates) == 0,
        (
            f"Duplicate dates="
            f"{len(duplicate_dates)}."
        ),
    )

    add_check(
        "No duplicate date/price records",
        len(duplicate_records) == 0,
        (
            f"Duplicate records="
            f"{len(duplicate_records)}."
        ),
    )

    # -----------------------------------------------------------------------
    # Newest / oldest observations
    # -----------------------------------------------------------------------

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
                f"BID={newest_observation.get('bidPrice')}."
            )
            if newest_observation
            else
            "No observations found."
        ),
    )

    add_check(
        "Oldest historical observation identified",
        oldest_observation is not None,
        (
            (
                f"{oldest_observation.get('date')} "
                f"BID={oldest_observation.get('bidPrice')}."
            )
            if oldest_observation
            else
            "No observations found."
        ),
    )

    # -----------------------------------------------------------------------
    # Historical data reaches inception
    # -----------------------------------------------------------------------

    inception_value = clean_text(
        fund_summary.get("inceptionDate")
    )

    if not inception_value:
        inception_value = clean_text(
            successful_summary.get(
                "inceptionDate"
            )
        )

    if not inception_value:
        inception_value = clean_text(
            fund_summary.get(
                "prudentialFundData",
                {},
            ).get("inceptionDate")
            if isinstance(
                fund_summary.get(
                    "prudentialFundData"
                ),
                dict,
            )
            else ""
        )

    inception_date = None

    try:

        if inception_value:

            inception_date = parse_date(
                inception_value
            )

    except ValueError:
        inception_date = None

    if (
        inception_date is not None
        and oldest_observation
    ):

        try:

            oldest_date = parse_date(
                oldest_observation["date"]
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
                f"Unable to validate inception. "
                f"Inception='{inception_value}'."
            ),
        )

    # -----------------------------------------------------------------------
    # Synthetic/generated data detection
    # -----------------------------------------------------------------------

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

    # -----------------------------------------------------------------------
    # Pagination validation
    # -----------------------------------------------------------------------

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
        isinstance(pages, list),
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
    full_page_violations = []
    final_page_valid = False

    for index, page in enumerate(
        pages
    ):

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
                        "page": page.get("page"),
                        "rowCount": row_count,
                    }
                )

        else:

            final_page_valid = (
                row_count > 0
                and row_count < PAGE_SIZE
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
        "Final pagination page is partial",
        final_page_valid,
        (
            (
                f"Final page rows="
                f"{pages[-1].get('rowCount')}."
            )
            if pages
            else
            "No pagination pages found."
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

    # -----------------------------------------------------------------------
    # Current Prudential BID vs latest PruAccess BID
    # -----------------------------------------------------------------------

    current_prudential_bid = None

    prudential_data = None

    if isinstance(
        prudential_fund,
        dict,
    ):

        prudential_data = (
            prudential_fund.get(
                "data"
            )
        )

    if not isinstance(
        prudential_data,
        dict,
    ):

        # Some versions may store the fund data directly.
        prudential_data = prudential_fund

    if isinstance(
        prudential_data,
        dict,
    ):

        current_prudential_bid = (
            prudential_data.get(
                "bidPrice"
            )
        )

    latest_pruaccess_bid = None

    if newest_observation:

        latest_pruaccess_bid = (
            newest_observation.get(
                "bidPrice"
            )
        )

    bid_match_passed = False
    bid_difference = None

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

    # -----------------------------------------------------------------------
    # Determine overall fund result
    # -----------------------------------------------------------------------

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
            "latestPruAccessDate": (
                newest_observation.get("date")
                if newest_observation
                else None
            ),
            "bidDifference": bid_difference,
        },
        "passed": passed,
    }

    return passed, result


# ---------------------------------------------------------------------------
# Main validation
# ---------------------------------------------------------------------------

def main() -> int:

    print("=" * 70)
    print(
        "PruAccess All-Fund Historical BID Price Validation"
    )
    print("=" * 70)

    try:

        # -------------------------------------------------------------------
        # Load master universe
        # -------------------------------------------------------------------

        excel_funds = read_excel_universe()

        print()
        print(
            f"Excel fund universe: "
            f"{len(excel_funds)}"
        )

        # -------------------------------------------------------------------
        # Load run summary
        # -------------------------------------------------------------------

        run_summary = load_json(
            RUN_SUMMARY_FILE
        )

        all_funds = load_json(
            ALL_FUNDS_FILE
        )

        all_bid_history = load_json(
            ALL_BID_HISTORY_FILE
        )

        # -------------------------------------------------------------------
        # Top-level extraction checks
        # -------------------------------------------------------------------

        expected_count = len(
            excel_funds
        )

        successful_count = run_summary.get(
            "successfulFundCount"
        )

        failed_count = run_summary.get(
            "failedFundCount"
        )

        print(
            f"Run successful funds: "
            f"{successful_count}"
        )

        print(
            f"Run failed funds: "
            f"{failed_count}"
        )

        print()

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
            expected_count == 67,
            (
                f"Found {expected_count} funds "
                f"in Funds Links.xlsm."
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
                f"Successful="
                f"{successful_count}; "
                f"expected="
                f"{expected_count}."
            ),
        )

        top_check(
            "No failed funds",
            failed_count == 0,
            (
                f"Failed funds="
                f"{failed_count}."
            ),
        )

        # -------------------------------------------------------------------
        # Check successful summary against Excel
        # -------------------------------------------------------------------

        successful_funds = run_summary.get(
            "successfulFunds"
        )

        if not isinstance(
            successful_funds,
            list,
        ):

            successful_funds = []

        successful_by_row = {}

        for fund in successful_funds:

            row = fund.get(
                "excelRow"
            )

            if row is not None:
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

        # -------------------------------------------------------------------
        # Validate each fund
        # -------------------------------------------------------------------

        print(
            "Validating individual funds..."
        )

        fund_results = []

        passed_funds = 0
        failed_funds = 0

        for index, excel_fund in enumerate(
            excel_funds,
            start=1,
        ):

            row = excel_fund[
                "excelRow"
            ]

            print(
                f"  [{index}/{len(excel_funds)}] "
                f"Excel row {row}: ",
                end="",
                flush=True,
            )

            successful_summary = (
                successful_by_row.get(row)
            )

            if successful_summary is None:

                result = {
                    "excelRow": row,
                    "fundName": "",
                    "fundIdentifier": "",
                    "fundCode": "",
                    "observationCount": 0,
                    "checks": [
                        {
                            "name": "Successful extraction exists",
                            "passed": False,
                            "details": (
                                "No successful extraction "
                                "exists for this Excel row."
                            ),
                        }
                    ],
                    "details": {},
                    "passed": False,
                }

                fund_results.append(
                    result
                )

                failed_funds += 1

                print(
                    "FAILED"
                )

                continue

            passed, result = validate_fund(
                excel_fund=excel_fund,
                successful_summary=successful_summary,
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

                failed_funds += 1

                print(
                    "FAILED"
                )

        # -------------------------------------------------------------------
        # Validate consolidated all_bid_history.json
        # -------------------------------------------------------------------

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

        top_check(
            "Consolidated BID history fund count",
            len(consolidated_funds)
            == expected_count,
            (
                f"Consolidated funds="
                f"{len(consolidated_funds)}; "
                f"expected="
                f"{expected_count}."
            ),
        )

        consolidated_total = 0

        consolidated_mismatches = []

        for result in fund_results:

            if not result.get(
                "passed"
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
                        "reason": "Missing consolidated fund.",
                    }
                )

                continue

            consolidated_count = consolidated.get(
                "observationCount"
            )

            actual_count = result[
                "observationCount"
            ]

            if consolidated_count != actual_count:

                consolidated_mismatches.append(
                    {
                        "fundIdentifier": identifier,
                        "reason": (
                            "Observation count mismatch."
                        ),
                        "consolidated": consolidated_count,
                        "individual": actual_count,
                    }
                )

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

            consolidated_total += len(
                consolidated_observations
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

        # -------------------------------------------------------------------
        # Validate consolidated total
        # -------------------------------------------------------------------

        expected_total = run_summary.get(
            "totalHistoricalBidObservations"
        )

        top_check(
            "Consolidated observation total",
            consolidated_total == expected_total,
            (
                f"Consolidated="
                f"{consolidated_total}; "
                f"run summary="
                f"{expected_total}."
            ),
        )

        # -------------------------------------------------------------------
        # Validate all_funds.json
        # -------------------------------------------------------------------

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
            == expected_count,
            (
                f"all_funds.json="
                f"{len(all_funds_list)}; "
                f"expected="
                f"{expected_count}."
            ),
        )

        # -------------------------------------------------------------------
        # Total individual observations
        # -------------------------------------------------------------------

        individual_total = sum(
            result.get(
                "observationCount",
                0,
            )
            for result in fund_results
            if result.get("passed")
        )

        top_check(
            "Individual observation total",
            individual_total
            == expected_total,
            (
                f"Individual="
                f"{individual_total}; "
                f"run summary="
                f"{expected_total}."
            ),
        )

        # -------------------------------------------------------------------
        # Overall result
        # -------------------------------------------------------------------

        all_top_checks_passed = all(
            check["passed"]
            for check in top_level_checks
        )

        overall_passed = (
            all_top_checks_passed
            and failed_funds == 0
            and passed_funds == expected_count
        )

        status = (
            "PASS"
            if overall_passed
            else "FAIL"
        )

        validation = {
            "status": status,

            "validatedAtUtc": (
                datetime.utcnow()
                .isoformat()
                + "Z"
            ),

            "excelFile": str(
                EXCEL_FILE
            ),

            "fundUniverseCount": expected_count,

            "passedFundCount": passed_funds,

            "failedFundCount": failed_funds,

            "totalHistoricalBidObservations": individual_total,

            "expectedTotalHistoricalBidObservations": (
                expected_total
            ),

            "topLevelChecks": top_level_checks,

            "fundResults": fund_results,

            "details": {
                "missingExcelRows": missing_excel_rows,
                "consolidatedMismatches": (
                    consolidated_mismatches
                ),
            },
        }

        save_json(
            VALIDATION_FILE,
            validation,
        )

        # -------------------------------------------------------------------
        # Console output
        # -------------------------------------------------------------------

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
            f"Funds passed: "
            f"{passed_funds}"
        )

        print(
            f"Funds failed: "
            f"{failed_funds}"
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
                else "FAIL"
            )

            print(
                f"  [{marker}] "
                f"{check['name']}: "
                f"{check['details']}"
            )

        if failed_funds:

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

                    if not check["passed"]:

                        print(
                            f"    - "
                            f"{check['name']}: "
                            f"{check['details']}"
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
                "ALL FUNDS PASSED VALIDATION."
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
            str(error)
        )

        return 1


if __name__ == "__main__":
    sys.exit(
        main()
    )
