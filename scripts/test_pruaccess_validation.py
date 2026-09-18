#!/usr/bin/env python3

"""
Validate the PruAccess historical BID extraction.

Test fund:
    PRULink ActiveInvest Portfolio - Balanced (Accumulation)

Input files:
    output_pruaccess/bid_history.json
    output_pruaccess/prudential_fund.json
    output_pruaccess/pagination.json

This script ONLY validates the extracted data.
It does not modify the source data.

Validation checks:
    - Correct fund
    - PruAccess source
    - BID price type
    - Expected observation count
    - Valid dates
    - Valid numeric BID prices
    - No missing fields
    - Dates strictly descending
    - No duplicate dates
    - No duplicate date/price records
    - Historical data reaches inception
    - Newest and oldest observations
    - No synthetic/interpolated fields
    - Prudential current BID vs PruAccess latest BID
    - Pagination count and row counts
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_DIR = Path("output_pruaccess")

BID_HISTORY_FILE = OUTPUT_DIR / "bid_history.json"
PRUDENTIAL_FILE = OUTPUT_DIR / "prudential_fund.json"
PAGINATION_FILE = OUTPUT_DIR / "pagination.json"
VALIDATION_FILE = OUTPUT_DIR / "validation.json"

EXPECTED_FUND_NAME = (
    "PRULink ActiveInvest Portfolio - Balanced (Accumulation)"
)

EXPECTED_PAGES = 62
EXPECTED_OBSERVATIONS = 1223

PruACCESS_DATE_FORMAT = "%d-%b-%Y"
PRUDENTIAL_DATE_FORMAT = "%d/%m/%Y"

# Prudential current bid is displayed to 4 decimal places.
# PruAccess historical bid contains 5 decimal places.
# Therefore this is a tolerance check, not exact equality.
CURRENT_BID_TOLERANCE = Decimal("0.0002")


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def load_json(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Required file not found: {path}"
        )

    return json.loads(
        path.read_text(encoding="utf-8")
    )


def save_json(path: Path, data) -> None:
    path.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def parse_pruaccess_date(value: str) -> datetime:
    return datetime.strptime(
        clean_text(value),
        PruACCESS_DATE_FORMAT,
    )


def parse_prudential_date(value: str) -> datetime:
    return datetime.strptime(
        clean_text(value),
        PRUDENTIAL_DATE_FORMAT,
    )


def parse_decimal(value) -> Decimal:
    text = clean_text(value)

    if not text:
        raise InvalidOperation(
            "Empty numeric value"
        )

    text = text.replace("$", "")
    text = text.replace(",", "")
    text = text.strip()

    number = Decimal(text)

    if not number.is_finite():
        raise InvalidOperation(
            "Number is not finite"
        )

    return number


def is_valid_price(value) -> bool:
    try:
        number = parse_decimal(value)

        return number > 0

    except (
        InvalidOperation,
        ValueError,
        TypeError,
    ):
        return False


def print_result(
    name: str,
    passed: bool,
    details: str = "",
) -> None:

    status = "PASS" if passed else "FAIL"

    print(
        f"[{status}] {name}"
    )

    if details:
        print(
            f"       {details}"
        )


# ============================================================
# MAIN VALIDATION
# ============================================================

def validate():
    print()
    print("=" * 70)
    print("PruAccess Historical BID Price Validation")
    print("=" * 70)
    print()

    # --------------------------------------------------------
    # Load files
    # --------------------------------------------------------

    bid_history = load_json(
        BID_HISTORY_FILE
    )

    prudential = load_json(
        PRUDENTIAL_FILE
    )

    pagination = load_json(
        PAGINATION_FILE
    )

    checks = []

    # --------------------------------------------------------
    # 1. Basic JSON structure
    # --------------------------------------------------------

    structure_ok = (
        isinstance(bid_history, dict)
        and isinstance(
            bid_history.get("observations"),
            list,
        )
    )

    checks.append({
        "name": "JSON structure",
        "passed": structure_ok,
        "details": (
            "observations list found."
            if structure_ok
            else "observations list missing."
        ),
    })

    if not structure_ok:
        return checks

    observations = bid_history["observations"]

    # --------------------------------------------------------
    # 2. Source
    # --------------------------------------------------------

    source = bid_history.get("source")

    source_ok = (
        source == "PruAccess"
    )

    checks.append({
        "name": "Historical source",
        "passed": source_ok,
        "details": (
            f"source={source!r}"
        ),
    })

    # --------------------------------------------------------
    # 3. BID only
    # --------------------------------------------------------

    price_type = bid_history.get(
        "priceType"
    )

    price_type_ok = (
        price_type == "BID"
    )

    checks.append({
        "name": "Historical price type",
        "passed": price_type_ok,
        "details": (
            f"priceType={price_type!r}"
        ),
    })

    # --------------------------------------------------------
    # 4. Fund identity
    # --------------------------------------------------------

    fund_name = clean_text(
        bid_history.get("fundName")
    )

    fund_ok = (
        fund_name == EXPECTED_FUND_NAME
    )

    checks.append({
        "name": "Fund identity",
        "passed": fund_ok,
        "details": (
            f"Fund={fund_name}"
        ),
    })

    # --------------------------------------------------------
    # 5. Observation count
    # --------------------------------------------------------

    actual_count = len(
        observations
    )

    count_ok = (
        actual_count
        == EXPECTED_OBSERVATIONS
    )

    checks.append({
        "name": "Historical observation count",
        "passed": count_ok,
        "details": (
            f"Found {actual_count}; "
            f"expected {EXPECTED_OBSERVATIONS}."
        ),
    })

    # --------------------------------------------------------
    # 6. Stored observation count
    # --------------------------------------------------------

    stored_count = bid_history.get(
        "observationCount"
    )

    stored_count_ok = (
        stored_count == actual_count
    )

    checks.append({
        "name": "Stored observationCount",
        "passed": stored_count_ok,
        "details": (
            f"Stored={stored_count}; "
            f"actual={actual_count}."
        ),
    })

    if not observations:
        checks.append({
            "name": "Historical observations available",
            "passed": False,
            "details": "No observations found.",
        })

        return checks

    # --------------------------------------------------------
    # 7. Validate every observation
    # --------------------------------------------------------

    invalid_dates = []
    invalid_prices = []
    missing_fields = []
    parsed_rows = []

    for index, observation in enumerate(
        observations,
        start=1,
    ):

        if not isinstance(
            observation,
            dict,
        ):
            missing_fields.append({
                "index": index,
                "reason": "Not an object",
            })
            continue

        date_value = observation.get(
            "date"
        )

        bid_value = observation.get(
            "bidPrice"
        )

        parsed_date = None

        # Date
        if not clean_text(date_value):

            missing_fields.append({
                "index": index,
                "reason": "Missing date",
            })

        else:

            try:

                parsed_date = (
                    parse_pruaccess_date(
                        date_value
                    )
                )

            except (
                ValueError,
                TypeError,
            ):

                invalid_dates.append({
                    "index": index,
                    "date": date_value,
                })

        # BID
        if not clean_text(bid_value):

            missing_fields.append({
                "index": index,
                "reason": "Missing bidPrice",
            })

        elif not is_valid_price(
            bid_value
        ):

            invalid_prices.append({
                "index": index,
                "bidPrice": bid_value,
            })

        if (
            parsed_date is not None
            and is_valid_price(bid_value)
        ):

            parsed_rows.append({
                "index": index,
                "dateText": clean_text(
                    date_value
                ),
                "date": parsed_date,
                "bidPriceText": clean_text(
                    bid_value
                ),
                "bidPrice": parse_decimal(
                    bid_value
                ),
            })

    dates_ok = (
        len(invalid_dates) == 0
    )

    checks.append({
        "name": "All dates valid",
        "passed": dates_ok,
        "details": (
            f"Invalid dates={len(invalid_dates)}."
        ),
    })

    prices_ok = (
        len(invalid_prices) == 0
    )

    checks.append({
        "name": "All BID prices numeric",
        "passed": prices_ok,
        "details": (
            f"Invalid prices={len(invalid_prices)}."
        ),
    })

    missing_ok = (
        len(missing_fields) == 0
    )

    checks.append({
        "name": "No missing observation fields",
        "passed": missing_ok,
        "details": (
            f"Missing/invalid records="
            f"{len(missing_fields)}."
        ),
    })

    # --------------------------------------------------------
    # 8. Date ordering
    # --------------------------------------------------------

    ordering_violations = []

    for previous, current in zip(
        parsed_rows,
        parsed_rows[1:],
    ):

        if current["date"] >= previous["date"]:

            ordering_violations.append({
                "previous": previous["dateText"],
                "current": current["dateText"],
            })

    ordering_ok = (
        len(ordering_violations) == 0
    )

    checks.append({
        "name": "Dates strictly descending",
        "passed": ordering_ok,
        "details": (
            f"Ordering violations="
            f"{len(ordering_violations)}."
        ),
    })

    # --------------------------------------------------------
    # 9. Duplicate dates
    # --------------------------------------------------------

    date_occurrences = {}

    for row in parsed_rows:

        date = row["dateText"]

        date_occurrences.setdefault(
            date,
            [],
        ).append(row)

    duplicate_dates = {
        date: rows
        for date, rows
        in date_occurrences.items()
        if len(rows) > 1
    }

    duplicate_dates_ok = (
        len(duplicate_dates) == 0
    )

    checks.append({
        "name": "No duplicate dates",
        "passed": duplicate_dates_ok,
        "details": (
            f"Duplicate dates="
            f"{len(duplicate_dates)}."
        ),
    })

    # --------------------------------------------------------
    # 10. Duplicate records
    # --------------------------------------------------------

    record_occurrences = {}

    for row in parsed_rows:

        key = (
            row["dateText"],
            row["bidPriceText"],
        )

        record_occurrences[key] = (
            record_occurrences.get(
                key,
                0,
            )
            + 1
        )

    duplicate_records = {
        key: count
        for key, count
        in record_occurrences.items()
        if count > 1
    }

    duplicate_records_ok = (
        len(duplicate_records) == 0
    )

    checks.append({
        "name": "No duplicate date/price records",
        "passed": duplicate_records_ok,
        "details": (
            f"Duplicate records="
            f"{len(duplicate_records)}."
        ),
    })

    # --------------------------------------------------------
    # 11. Inception date coverage
    # --------------------------------------------------------

    inception_raw = clean_text(
        prudential.get(
            "inceptionDate"
        )
    )

    inception_ok = False
    inception_details = ""

    try:

        inception_date = (
            parse_prudential_date(
                inception_raw
            )
        )

        oldest_row = min(
            parsed_rows,
            key=lambda row: row["date"],
        )

        inception_ok = (
            oldest_row["date"].date()
            <= inception_date.date()
        )

        inception_details = (
            f"Inception="
            f"{inception_raw}; "
            f"oldest observation="
            f"{oldest_row['dateText']}."
        )

    except Exception as exc:

        inception_details = (
            f"Unable to validate: {exc}"
        )

    checks.append({
        "name": "Historical data reaches inception",
        "passed": inception_ok,
        "details": inception_details,
    })

    # --------------------------------------------------------
    # 12. Newest observation
    # --------------------------------------------------------

    newest_row = max(
        parsed_rows,
        key=lambda row: row["date"],
    )

    checks.append({
        "name": "Newest historical observation identified",
        "passed": True,
        "details": (
            f"{newest_row['dateText']} "
            f"BID={newest_row['bidPriceText']}."
        ),
    })

    # --------------------------------------------------------
    # 13. Oldest observation
    # --------------------------------------------------------

    oldest_row = min(
        parsed_rows,
        key=lambda row: row["date"],
    )

    checks.append({
        "name": "Oldest historical observation identified",
        "passed": True,
        "details": (
            f"{oldest_row['dateText']} "
            f"BID={oldest_row['bidPriceText']}."
        ),
    })

    # --------------------------------------------------------
    # 14. Synthetic data detection
    # --------------------------------------------------------

    forbidden_fields = {
        "synthetic",
        "estimated",
        "estimate",
        "interpolated",
        "interpolation",
        "simulated",
        "simulation",
        "generated",
        "carryforward",
        "carriedforward",
        "forwardfilled",
    }

    suspicious_fields = []

    for index, observation in enumerate(
        observations,
        start=1,
    ):

        if not isinstance(
            observation,
            dict,
        ):
            continue

        for field in observation.keys():

            normalized = (
                clean_text(field)
                .replace("_", "")
                .replace("-", "")
                .lower()
            )

            if normalized in forbidden_fields:

                suspicious_fields.append({
                    "index": index,
                    "field": field,
                })

    synthetic_ok = (
        len(suspicious_fields) == 0
    )

    checks.append({
        "name": "No synthetic/interpolated fields",
        "passed": synthetic_ok,
        "details": (
            "No generated-data fields detected."
            if synthetic_ok
            else (
                f"Suspicious fields="
                f"{len(suspicious_fields)}."
            )
        ),
    })

    # --------------------------------------------------------
    # 15. Current Prudential BID comparison
    # --------------------------------------------------------

    prudential_bid_raw = clean_text(
        prudential.get(
            "bidPrice"
        )
    )

    current_bid_ok = False
    current_bid_details = ""

    try:

        prudential_bid = parse_decimal(
            prudential_bid_raw
        )

        latest_historical_bid = (
            newest_row["bidPrice"]
        )

        difference = abs(
            prudential_bid
            - latest_historical_bid
        )

        current_bid_ok = (
            difference
            <= CURRENT_BID_TOLERANCE
        )

        current_bid_details = (
            f"Prudential current BID="
            f"{prudential_bid}; "
            f"PruAccess latest BID="
            f"{latest_historical_bid}; "
            f"difference={difference}; "
            f"tolerance="
            f"{CURRENT_BID_TOLERANCE}."
        )

    except Exception as exc:

        current_bid_details = (
            f"Comparison failed: {exc}"
        )

    checks.append({
        "name": "Current Prudential BID matches latest PruAccess BID",
        "passed": current_bid_ok,
        "details": current_bid_details,
    })

    # --------------------------------------------------------
    # 16. Pagination validation
    # --------------------------------------------------------

    pagination_ok = isinstance(
        pagination,
        list,
    )

    checks.append({
        "name": "Pagination JSON structure",
        "passed": pagination_ok,
        "details": (
            f"Pages found="
            f"{len(pagination) if pagination_ok else 0}."
        ),
    })

    pagination_page_count_ok = False
    pagination_rows_ok = False
    final_page_ok = False

    if pagination_ok:

        pagination_page_count_ok = (
            len(pagination)
            == EXPECTED_PAGES
        )

        checks.append({
            "name": "Pagination page count",
            "passed": pagination_page_count_ok,
            "details": (
                f"Found {len(pagination)}; "
                f"expected {EXPECTED_PAGES}."
            ),
        })

        row_counts = []

        for item in pagination:

            if isinstance(item, dict):

                row_counts.append(
                    item.get("rowCount")
                )

        if len(row_counts) == EXPECTED_PAGES:

            full_page_counts = (
                row_counts[:-1]
            )

            pagination_rows_ok = all(
                count == 20
                for count
                in full_page_counts
            )

            final_page_ok = (
                row_counts[-1] < 20
            )

        checks.append({
            "name": "Pagination full pages contain 20 rows",
            "passed": pagination_rows_ok,
            "details": (
                f"Full pages checked="
                f"{max(len(row_counts) - 1, 0)}."
            ),
        })

        checks.append({
            "name": "Final pagination page is partial",
            "passed": final_page_ok,
            "details": (
                f"Final page rows="
                f"{row_counts[-1] if row_counts else 'N/A'}."
            ),
        })

    # --------------------------------------------------------
    # 17. Pagination row total
    # --------------------------------------------------------

    if pagination_ok:

        pagination_total = sum(
            item.get("rowCount", 0)
            for item in pagination
            if isinstance(item, dict)
            and isinstance(
                item.get("rowCount"),
                int,
            )
        )

        pagination_total_ok = (
            pagination_total
            == EXPECTED_OBSERVATIONS
        )

        checks.append({
            "name": "Pagination row total matches observations",
            "passed": pagination_total_ok,
            "details": (
                f"Pagination rows="
                f"{pagination_total}; "
                f"observations="
                f"{actual_count}."
            ),
        })

    else:

        pagination_total_ok = False

        checks.append({
            "name": "Pagination row total matches observations",
            "passed": False,
            "details": (
                "Pagination data unavailable."
            ),
        })

    # --------------------------------------------------------
    # OVERALL RESULT
    # --------------------------------------------------------

    overall_passed = all(
        check["passed"]
        for check in checks
    )

    # --------------------------------------------------------
    # Build validation report
    # --------------------------------------------------------

    validation_result = {
        "status": (
            "PASS"
            if overall_passed
            else "FAIL"
        ),
        "validatedAtUtc": (
            datetime.utcnow().isoformat()
            + "Z"
        ),
        "fundName": fund_name,
        "source": source,
        "priceType": price_type,
        "expectedObservationCount": (
            EXPECTED_OBSERVATIONS
        ),
        "actualObservationCount": (
            actual_count
        ),
        "expectedPages": EXPECTED_PAGES,
        "actualPages": (
            len(pagination)
            if isinstance(
                pagination,
                list,
            )
            else None
        ),
        "currentPrudentialBid": (
            prudential_bid_raw
        ),
        "latestPruAccessBid": (
            newest_row["bidPriceText"]
        ),
        "latestPruAccessDate": (
            newest_row["dateText"]
        ),
        "oldestPruAccessBid": (
            oldest_row["bidPriceText"]
        ),
        "oldestPruAccessDate": (
            oldest_row["dateText"]
        ),
        "checks": checks,
        "details": {
            "invalidDates": invalid_dates,
            "invalidPrices": invalid_prices,
            "missingFields": missing_fields,
            "orderingViolations": ordering_violations,
            "duplicateDates": {
                date: [
                    row["index"]
                    for row in rows
                ]
                for date, rows
                in duplicate_dates.items()
            },
            "duplicateRecords": {
                f"{key[0]} | {key[1]}": count
                for key, count
                in duplicate_records.items()
            },
            "suspiciousGeneratedFields": (
                suspicious_fields
            ),
        },
    }

    save_json(
        VALIDATION_FILE,
        validation_result,
    )

    # --------------------------------------------------------
    # Print results
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("VALIDATION RESULTS")
    print("=" * 70)
    print()

    for check in checks:

        print_result(
            check["name"],
            check["passed"],
            check["details"],
        )

    print()
    print("=" * 70)

    if overall_passed:

        print(
            "OVERALL RESULT: PASS"
        )

    else:

        print(
            "OVERALL RESULT: FAIL"
        )

    print("=" * 70)

    print()
    print("SUMMARY")
    print("-" * 70)

    print(
        f"Fund:                  {fund_name}"
    )

    print(
        f"Observations:          {actual_count}"
    )

    print(
        f"Expected observations: {EXPECTED_OBSERVATIONS}"
    )

    print(
        f"Pages:                 "
        f"{len(pagination) if isinstance(pagination, list) else 'N/A'}"
    )

    print(
        f"Newest date:           "
        f"{newest_row['dateText']}"
    )

    print(
        f"Newest BID:            "
        f"{newest_row['bidPriceText']}"
    )

    print(
        f"Oldest date:           "
        f"{oldest_row['dateText']}"
    )

    print(
        f"Oldest BID:            "
        f"{oldest_row['bidPriceText']}"
    )

    print(
        f"Prudential current BID:"
        f" {prudential_bid_raw}"
    )

    print()
    print(
        f"Validation report:"
        f" {VALIDATION_FILE}"
    )

    print()

    return overall_passed


# ============================================================
# ENTRY POINT
# ============================================================

def main():

    try:

        passed = validate()

    except Exception as exc:

        print()
        print("=" * 70)
        print("VALIDATION ERROR")
        print("=" * 70)
        print()
        print(str(exc))
        print()

        raise SystemExit(1)

    if not passed:

        raise SystemExit(1)


if __name__ == "__main__":
    main()


### Then run

From the **root of `vgrat-fmms`**:

```bash
python scripts/test_pruaccess_validation.py
```

It expects these files that your extraction already created:

```text
output_pruaccess/
bid_history.json
prudential_fund.json
pagination.json
```

It will create:

```text
output_pruaccess/
validation.json
```

The important thing I want to see next is the bottom section:

```text
======================================================================
VALIDATION RESULTS
======================================================================

[PASS] ...
[PASS] ...
...

======================================================================
OVERALL RESULT: PASS
======================================================================


Run it and paste that output here.
