
#!/usr/bin/env python3

"""
Validate PruAccess historical BID price extraction.

Test fund:
    PRULink ActiveInvest Portfolio - Balanced (Accumulation)

Expected source:
    output_pruaccess/bid_history.json

Validation checks:
    1. Historical observation count
    2. Valid dates
    3. Date ordering
    4. Duplicate dates
    5. Duplicate date/price records
    6. Numeric BID prices
    7. Inception-date coverage
    8. Latest available observation
    9. No synthetic/interpolated values
   10. Prudential current BID vs latest PruAccess BID

Important:
- This script does NOT modify bid_history.json.
- It does NOT create missing data.
- It does NOT interpolate prices.
- It does NOT calculate replacement prices.
- It treats PruAccess observations as actual observations only.
"""

from __future__ import annotations

import json
import math
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path


# ============================================================
# Configuration
# ============================================================

OUTPUT_DIR = Path("output_pruaccess")

BID_HISTORY_FILE = OUTPUT_DIR / "bid_history.json"
PRUDENTIAL_FILE = OUTPUT_DIR / "prudential_fund.json"

VALIDATION_FILE = OUTPUT_DIR / "validation.json"

EXPECTED_FUND_NAME = (
    "PRULink ActiveInvest Portfolio - Balanced (Accumulation)"
)

EXPECTED_PAGES = 62
EXPECTED_OBSERVATIONS = 1223

DATE_FORMAT = "%d-%b-%Y"

# Current Prudential bid is displayed to 4 decimal places,
# while PruAccess historical bid prices can contain 5 decimals.
# Therefore, this is a comparison/diagnostic tolerance rather
# than an exact-equality requirement.
CURRENT_BID_TOLERANCE = Decimal("0.0002")


# ============================================================
# Helpers
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def parse_date(value: str) -> datetime:
    return datetime.strptime(
        clean_text(value),
        DATE_FORMAT,
    )


def parse_decimal(value) -> Decimal:
    """
    Parse a price safely.

    Accepts:
        1.11916
        "$1.11916"
        "1,119.16"

    Does not accept:
        -
        blank
        arbitrary text
    """

    text = clean_text(value)

    if not text:
        raise InvalidOperation("Empty numeric value")

    text = text.replace("$", "")
    text = text.replace(",", "")
    text = text.strip()

    return Decimal(text)


def is_numeric_price(value) -> bool:
    try:
        number = parse_decimal(value)

        if not number.is_finite():
            return False

        if number <= 0:
            return False

        return True

    except (InvalidOperation, ValueError, TypeError):
        return False


def print_check(name: str, passed: bool, details: str = "") -> None:
    status = "PASS" if passed else "FAIL"

    print(f"[{status}] {name}")

    if details:
        print(f"       {details}")


def load_json(path: Path):
    if not path.exists():
        raise FileNotFoundError(
            f"Required file not found: {path}"
        )

    return json.loads(
        path.read_text(encoding="utf-8")
    )


# ============================================================
# Validation
# ============================================================

def validate_bid_history():
    print("=" * 70)
    print("PruAccess Historical BID Price Validation")
    print("=" * 70)

    print()
    print(f"BID history file: {BID_HISTORY_FILE}")
    print(f"Prudential file:  {PRUDENTIAL_FILE}")
    print()

    bid_history = load_json(BID_HISTORY_FILE)
    prudential = load_json(PRUDENTIAL_FILE)

    checks = []

    # --------------------------------------------------------
    # Basic structure
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
            "bid_history.json contains an observations list."
            if structure_ok
            else "Missing observations list."
        ),
    })

    if not structure_ok:
        return checks

    observations = bid_history["observations"]

    # --------------------------------------------------------
    # Source
    # --------------------------------------------------------

    source_value = bid_history.get("source")

    source_ok = source_value == "PruAccess"

    checks.append({
        "name": "Historical source",
        "passed": source_ok,
        "details": (
            f"source={source_value!r}"
        ),
    })

    # --------------------------------------------------------
    # Price type
    # --------------------------------------------------------

    price_type = bid_history.get("priceType")

    price_type_ok = price_type == "BID"

    checks.append({
        "name": "Historical price type",
        "passed": price_type_ok,
        "details": (
            f"priceType={price_type!r}"
        ),
    })

    # --------------------------------------------------------
    # Fund name
    # --------------------------------------------------------

    fund_name = clean_text(
        bid_history.get("fundName")
    )

    fund_name_ok = (
        fund_name == EXPECTED_FUND_NAME
    )

    checks.append({
        "name": "Fund identity",
        "passed": fund_name_ok,
        "details": (
            f"Fund={fund_name}"
        ),
    })

    # --------------------------------------------------------
    # Observation count
    # --------------------------------------------------------

    observation_count = len(observations)

    count_ok = (
        observation_count == EXPECTED_OBSERVATIONS
    )

    checks.append({
        "name": "Historical observation count",
        "passed": count_ok,
        "details": (
            f"Found {observation_count}; "
            f"expected {EXPECTED_OBSERVATIONS}."
        ),
    })

    # --------------------------------------------------------
    # Stored observationCount
    # --------------------------------------------------------

    stored_count = bid_history.get(
        "observationCount"
    )

    stored_count_ok = (
        stored_count == observation_count
    )

    checks.append({
        "name": "Stored observationCount",
        "passed": stored_count_ok,
        "details": (
            f"JSON observationCount={stored_count}; "
            f"actual list length={observation_count}."
        ),
    })

    # --------------------------------------------------------
    # Empty data check
    # --------------------------------------------------------

    if not observations:
        checks.append({
            "name": "Historical observations available",
            "passed": False,
            "details": "Observation list is empty.",
        })

        return checks

    checks.append({
        "name": "Historical observations available",
        "passed": True,
        "details": (
            f"{len(observations)} observations available."
        ),
    })

    # --------------------------------------------------------
    # Validate every observation
    # --------------------------------------------------------

    invalid_dates = []
    invalid_prices = []
    missing_fields = []

    parsed_rows = []

    for index, observation in enumerate(observations, start=1):

        if not isinstance(observation, dict):
            missing_fields.append({
                "index": index,
                "reason": "Observation is not an object",
            })
            continue

        date_value = observation.get("date")
        bid_value = observation.get("bidPrice")

        if not clean_text(date_value):
            missing_fields.append({
                "index": index,
                "reason": "Missing date",
            })
        else:
            try:
                parsed_date = parse_date(date_value)
            except (ValueError, TypeError):
                invalid_dates.append({
                    "index": index,
                    "date": date_value,
                })
                parsed_date = None

        if not clean_text(bid_value):
            missing_fields.append({
                "index": index,
                "reason": "Missing bidPrice",
            })
        elif not is_numeric_price(bid_value):
            invalid_prices.append({
                "index": index,
                "bidPrice": bid_value,
            })

        if (
            clean_text(date_value)
            and parsed_date is not None
            and is_numeric_price(bid_value)
        ):
            parsed_rows.append({
                "index": index,
                "dateText": clean_text(date_value),
                "date": parsed_date,
                "bidPriceText": clean_text(bid_value),
                "bidPrice": parse_decimal(bid_value),
            })

    dates_ok = len(invalid_dates) == 0

    checks.append({
        "name": "All dates valid",
        "passed": dates_ok,
        "details": (
            f"Invalid dates: {len(invalid_dates)}."
        ),
    })

    prices_ok = len(invalid_prices) == 0

    checks.append({
        "name": "All BID prices numeric",
        "passed": prices_ok,
        "details": (
            f"Invalid prices: {len(invalid_prices)}."
        ),
    })

    missing_ok = len(missing_fields) == 0

    checks.append({
        "name": "No missing observation fields",
        "passed": missing_ok,
        "details": (
            f"Invalid/missing records: "
            f"{len(missing_fields)}."
        ),
    })

    # --------------------------------------------------------
    # Date ordering
    #
    # Expected:
    # newest → oldest
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

    ordering_ok = len(ordering_violations) == 0

    checks.append({
        "name": "Dates strictly descending",
        "passed": ordering_ok,
        "details": (
            f"Ordering violations: "
            f"{len(ordering_violations)}."
        ),
    })

    # --------------------------------------------------------
    # Duplicate dates
    # --------------------------------------------------------

    date_map = {}

    for row in parsed_rows:
        date_map.setdefault(
            row["dateText"],
            []
        ).append(row)

    duplicate_dates = {
        date: rows
        for date, rows in date_map.items()
        if len(rows) > 1
    }

    duplicate_dates_ok = len(duplicate_dates) == 0

    checks.append({
        "name": "No duplicate dates",
        "passed": duplicate_dates_ok,
        "details": (
            f"Duplicate date values: "
            f"{len(duplicate_dates)}."
        ),
    })

    # --------------------------------------------------------
    # Duplicate date/price records
    # --------------------------------------------------------

    record_counts = {}

    for row in parsed_rows:
        key = (
            row["dateText"],
            row["bidPriceText"],
        )

        record_counts[key] = (
            record_counts.get(key, 0) + 1
        )

    duplicate_records = {
        key: count
        for key, count in record_counts.items()
        if count > 1
    }

    duplicate_records_ok = (
        len(duplicate_records) == 0
    )

    checks.append({
        "name": "No duplicate date/price records",
        "passed": duplicate_records_ok,
        "details": (
            f"Duplicate records: "
            f"{len(duplicate_records)}."
        ),
    })

    # --------------------------------------------------------
    # Inception coverage
    # --------------------------------------------------------

    inception_raw = clean_text(
        prudential.get("inceptionDate")
    )

    inception_ok = True
    inception_details = ""

    try:
        inception_date = datetime.strptime(
            inception_raw,
            "%d/%m/%Y",
        )

        actual_oldest = min(
            row["date"]
            for row in parsed_rows
        )

        inception_ok = (
            actual_oldest.date()
            <= inception_date.date()
        )

        inception_details = (
            f"Prudential inception={inception_raw}; "
            f"oldest PruAccess observation="
            f"{actual_oldest.strftime(DATE_FORMAT)}."
        )

    except Exception as exc:
        inception_ok = False
        inception_details = (
            f"Could not validate inception coverage: {exc}"
        )

    checks.append({
        "name": "Historical data reaches inception",
        "passed": inception_ok,
        "details": inception_details,
    })

    # --------------------------------------------------------
    # Newest observation
    # --------------------------------------------------------

    newest_row = max(
        parsed_rows,
        key=lambda row: row["date"],
    )

    newest_ok = True

    newest_details = (
        f"Newest observation="
        f"{newest_row['dateText']} "
        f"BID={newest_row['bidPriceText']}."
    )

    checks.append({
        "name": "Newest historical observation identified",
        "passed": newest_ok,
        "details": newest_details,
    })

    # --------------------------------------------------------
    # Oldest observation
    # --------------------------------------------------------

    oldest_row = min(
        parsed_rows,
        key=lambda row: row["date"],
    )

    oldest_ok = True

    oldest_details = (
        f"Oldest observation="
        f"{oldest_row['dateText']} "
        f"BID={oldest_row['bidPriceText']}."
    )

    checks.append({
        "name": "Oldest historical observation identified",
        "passed": oldest_ok,
        "details": oldest_details,
    })

    # --------------------------------------------------------
    # No synthetic/interpolated data
    #
    # The extractor itself stores source observations only.
    # We therefore inspect the records for fields that would
    # indicate generated values.
    # --------------------------------------------------------

    suspicious_generated_fields = []

    forbidden_generation_fields = {
        "synthetic",
        "estimated",
        "estimate",
        "interpolated",
        "interpolation",
        "simulated",
        "simulation",
        "generated",
        "carryForward",
        "carriedForward",
        "forwardFilled",
    }

    for index, observation in enumerate(
        observations,
        start=1,
    ):
        if not isinstance(observation, dict):
            continue

        found = []

        for key in observation.keys():
            normalized_key = (
                clean_text(key)
                .replace("_", "")
                .replace("-", "")
                .lower()
            )

            for forbidden in forbidden_generation_fields:
                normalized_forbidden = (
                    forbidden
                    .replace("_", "")
                    .replace("-", "")
                    .lower()
                )

                if normalized_key == normalized_forbidden:
                    found.append(key)

        if found:
            suspicious_generated_fields.append({
                "index": index,
                "fields": found,
            })

    no_generated_fields_ok = (
        len(suspicious_generated_fields) == 0
    )

    checks.append({
        "name": "No synthetic/interpolated fields",
        "passed": no_generated_fields_ok,
        "details": (
            "No generated-data fields detected."
            if no_generated_fields_ok
            else (
                f"Potential generated-data fields found "
                f"in {len(suspicious_generated_fields)} "
                f"observations."
            )
        ),
    })

    # --------------------------------------------------------
    # Prudential current BID vs PruAccess newest BID
    # --------------------------------------------------------

    prudential_bid_raw = clean_text(
        prudential.get("bidPrice")
    )

    current_bid_ok = False
    current_bid_details = ""

    try:
        prudential_bid = parse_decimal(
            prudential_bid_raw
        )

        historical_latest_bid = (
            newest_row["bidPrice"]
        )

        difference = abs(
            prudential_bid
            - historical_latest_bid
        )

        current_bid_ok = (
            difference <= CURRENT_BID_TOLERANCE
        )

        current_bid_details = (
            f"Prudential current BID="
            f"{prudential_bid}; "
            f"PruAccess latest BID="
            f"{historical_latest_bid}; "
            f"difference={difference}; "
            f"tolerance={CURRENT_BID_TOLERANCE}."
        )

    except Exception as exc:
        current_bid_details = (
            f"Could not compare current BID: {exc}"
        )

    checks.append({
        "name": "Current Prudential BID matches latest PruAccess BID",
        "passed": current_bid_ok,
        "details": current_bid_details,
    })

    # --------------------------------------------------------
    # Expected pagination
    #
    # pagination.json is optional here. If available, verify
    # the 62-page result independently.
    # --------------------------------------------------------

    pagination_file = OUTPUT_DIR / "pagination.json"

    pagination_checks = []

    if pagination_file.exists():

        try:
            pagination = load_json(
                pagination_file
            )

            if isinstance(pagination, list):

                page_count = len(pagination)

                pages_ok = (
                    page_count
                    == EXPECTED_PAGES
                )

                pagination_checks.append({
                    "name": "Pagination page count",
                    "passed": pages_ok,
                    "details": (
                        f"Found {page_count} pages; "
                        f"expected {EXPECTED_PAGES}."
                    ),
                })

                row_counts = [
                    item.get("rowCount")
                    for item in pagination
                    if isinstance(item, dict)
                ]

                expected_full_pages = (
                    row_counts[:-1]
                    if row_counts
                    else []
                )

                full_pages_ok = all(
                    count == 20
                    for count in expected_full_pages
                )

                pagination_checks.append({
                    "name": "Full pagination pages contain 20 rows",
                    "passed": full_pages_ok,
                    "details": (
                        f"Checked {len(expected_full_pages)} "
                        f"non-final pages."
                    ),
                })

                final_page_ok = False

                if row_counts:
                    final_page_ok = (
                        row_counts[-1] < 20
                    )

                pagination_checks.append({
                    "name": "Final pagination page is partial",
                    "passed": final_page_ok,
                    "details": (
                        f"Final page row count="
                        f"{row_counts[-1] if row_counts else 'N/A'}."
                    ),
                })

            else:
                pagination_checks.append({
                    "name": "Pagination JSON structure",
                    "passed": False,
                    "details": (
                        "pagination.json is not a list."
                    ),
                })

        except Exception as exc:
            pagination_checks.append({
                "name": "Pagination JSON readable",
                "passed": False,
                "details": str(exc),
            })

    else:
        pagination_checks.append({
            "name": "Pagination diagnostics available",
            "passed": False,
            "details": (
                "pagination.json was not found."
            ),
        })

    checks.extend(pagination_checks)

    # --------------------------------------------------------
    # Overall result
    # --------------------------------------------------------

    overall_passed = all(
        check["passed"]
        for check in checks
    )

    # --------------------------------------------------------
    # Detailed result
    # --------------------------------------------------------

    validation_result = {
        "status": (
            "PASS"
            if overall_passed
            else "FAIL"
        ),
        "validatedAtUtc": datetime.utcnow().isoformat()
        + "Z",
        "fundName": fund_name,
        "source": source_value,
        "priceType": price_type,
        "expectedObservationCount": EXPECTED_OBSERVATIONS,
        "actualObservationCount": observation_count,
        "expectedPages": EXPECTED_PAGES,
        "currentPrudentialBid": prudential_bid_raw,
        "latestPruAccessBid": (
            newest_row["bidPriceText"]
            if newest_row
            else None
        ),
        "latestPruAccessDate": (
            newest_row["dateText"]
            if newest_row
            else None
        ),
        "oldestPruAccessDate": (
            oldest_row["dateText"]
            if oldest_row
            else None
        ),
        "checks": checks,
        "details": {
            "invalidDates": invalid_dates,
            "invalidPrices": invalid_prices,
            "missingFields": missing_fields,
            "orderingViolations": ordering_violations,
            "duplicateDates": {
                key: [
                    row["index"]
                    for row in rows
                ]
                for key, rows in duplicate_dates.items()
            },
            "duplicateRecords": {
                f"{key[0]} | {key[1]}": count
                for key, count
                in duplicate_records.items()
            },
            "suspiciousGeneratedFields": (
                suspicious_generated_fields
            ),
        },
    }

    save_json(
        VALIDATION_FILE,
        validation_result,
    )

    # --------------------------------------------------------
    # Print result
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("VALIDATION RESULTS")
    print("=" * 70)

    print()

    for check in checks:
        print_check(
            check["name"],
            check["passed"],
            check["details"],
        )

    print()
    print("=" * 70)

    if overall_passed:
        print("OVERALL RESULT: PASS")
    else:
        print("OVERALL RESULT: FAIL")

    print("=" * 70)

    print()
    print("Key results:")

    print(
        f"  Fund:                 {fund_name}"
    )

    print(
        f"  Observations:         "
        f"{observation_count}"
    )

    print(
        f"  Expected observations:"
        f" {EXPECTED_OBSERVATIONS}"
    )

    print(
        f"  Newest PruAccess date: "
        f"{newest_row['dateText']}"
    )

    print(
        f"  Newest PruAccess BID:  "
        f"{newest_row['bidPriceText']}"
    )

    print(
        f"  Oldest PruAccess date: "
        f"{oldest_row['dateText']}"
    )

    print(
        f"  Oldest PruAccess BID:  "
        f"{oldest_row['bidPriceText']}"
    )

    print(
        f"  Prudential current BID:"
        f" {prudential_bid_raw}"
    )

    print()
    print(
        f"Validation report written to:"
        f" {VALIDATION_FILE}"
    )

    print()

    return overall_passed


# ============================================================
# Main
# ============================================================

def main():
    try:
        passed = validate_bid_history()

    except Exception as exc:
        print()
        print("=" * 70)
        print("VALIDATION ERROR")
        print("=" * 70)
        print()
        print(str(exc))
        print()
        raise

    if not passed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
```

### Run it

From the repository root:

```bash
python scripts/test_pruaccess_validation.py
```

You should ideally get:

```text
[PASS] JSON structure
[PASS] Historical source
[PASS] Historical price type
[PASS] Fund identity
[PASS] Historical observation count
[PASS] Stored observationCount
[PASS] Historical observations available
[PASS] All dates valid
[PASS] All BID prices numeric
[PASS] No missing observation fields
[PASS] Dates strictly descending
[PASS] No duplicate dates
[PASS] No duplicate date/price records
[PASS] Historical data reaches inception
[PASS] Newest historical observation identified
[PASS] Oldest historical observation identified
[PASS] No synthetic/interpolated fields
[PASS] Current Prudential BID matches latest PruAccess BID
[PASS] Pagination page count
...
OVERALL RESULT: PASS
```

It will also create:

```text
output_pruaccess/
└── validation.json


**Run this first.** Send me the final `VALIDATION RESULTS` section, especially if anything shows `FAIL`.
