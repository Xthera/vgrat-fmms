#!/usr/bin/env python3

"""
Validation suite for PruAccess historical BID extraction.

MASTER SOURCE
=============

    Funds Links.xlsm

Validation compares the extracted data against:

    1. Funds Links.xlsm
    2. output_pruaccess/run_summary.json
    3. output_pruaccess/all_funds.json
    4. output_pruaccess/all_bid_history.json
    5. individual fund JSON files

IMPORTANT RULES
===============

1. Excel Column A dynamically defines the fund universe.
2. Excel Column B defines the exact PruAccess lookup name.
3. No hardcoded fund count.
4. Prudential fund identity is authoritative.
5. Historical prices must come from PruAccess.
6. Historical price type must be BID.
7. Fund Price Type must not be manipulated.
8. No synthetic data.
9. No estimated data.
10. No interpolation.
11. No fabricated values.
12. No carry-forward values.
13. Every pagination page must be present.
14. Every page must have the expected row count.
15. Historical dates must be strictly descending.
16. Duplicate dates are errors.
17. Duplicate date/price observations are errors.
18. Historical data must reach the Prudential inception date.
19. Latest PruAccess BID must agree with current Prudential BID
    within the configured tolerance.
20. Missing funds cause validation failure.
21. Failed extraction funds cause validation failure.
22. The validator must never crash merely because a JSON field
    contains a list where a dictionary was expected.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path(
    "Funds Links.xlsm"
)

OUTPUT_DIR = Path(
    "output_pruaccess"
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

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR
    / "funds"
)

VALIDATION_FILE = (
    OUTPUT_DIR
    / "validation.json"
)

PAGE_SIZE = 20

PRICE_TOLERANCE = 0.0002


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def normalize_text(value) -> str:

    value = clean_text(
        value
    ).lower()

    value = value.replace(
        "–",
        "-"
    )

    value = value.replace(
        "—",
        "-"
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip()


def load_json(
    path: Path,
):

    if not path.exists():

        raise FileNotFoundError(
            f"Required JSON file not found: {path}"
        )

    try:

        return json.loads(
            path.read_text(
                encoding="utf-8"
            )
        )

    except Exception as error:

        raise RuntimeError(
            f"Could not parse JSON file "
            f"{path}: {error}"
        ) from error


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


def parse_price(
    value,
) -> float:

    text = clean_text(
        value
    )

    text = text.replace(
        "$",
        ""
    )

    text = text.replace(
        ",",
        ""
    )

    if not re.fullmatch(
        r"\d+(?:\.\d+)?",
        text,
    ):

        raise ValueError(
            f"Invalid numeric price: {value}"
        )

    result = float(
        text
    )

    if result <= 0:

        raise ValueError(
            f"Price must be greater than zero: "
            f"{value}"
        )

    return result


def parse_date(
    value: str,
) -> datetime:

    value = clean_text(
        value
    )

    formats = [
        "%d-%b-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                value,
                fmt,
            )

        except ValueError:

            continue

    raise ValueError(
        f"Unsupported date format: {value}"
    )


def get_fund_directory(
    excel_row: int,
) -> Path | None:
    """
    Locate the individual fund directory by Excel row.

    Expected form:

        output_pruaccess/funds/2_D39F
        output_pruaccess/funds/31_<identifier>
        etc.

    The identifier itself is not assumed.
    """

    if not FUNDS_OUTPUT_DIR.exists():

        return None

    prefix = f"{excel_row}_"

    matches = [
        path
        for path in FUNDS_OUTPUT_DIR.iterdir()
        if path.is_dir()
        and path.name.startswith(prefix)
        and not path.name.endswith("_failed")
    ]

    if not matches:

        return None

    if len(matches) > 1:

        raise RuntimeError(
            f"Multiple fund directories found "
            f"for Excel row {excel_row}: "
            f"{matches}"
        )

    return matches[0]


# ============================================================
# EXCEL UNIVERSE
# ============================================================

def read_excel_funds() -> list[dict]:

    print(
        "\nReading Excel fund universe..."
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

    workbook.close()

    if not funds:

        raise RuntimeError(
            "Excel Column A contains no populated "
            "Prudential fund URLs."
        )

    return funds


# ============================================================
# STRUCTURE NORMALIZATION
# ============================================================

FUND_IDENTITY_KEYS = {
    "excelrow",
    "excel_row",
    "prudentialfundname",
    "fundname",
    "fundidentifier",
    "fundcode",
    "pruaccessfundid",
    "pruaccessfundname",
    "pruaccessfundname",
    "prudentialurl",
}


def is_probable_fund_record(
    value,
) -> bool:
    """
    Determine whether a dictionary looks like a fund record.

    This prevents observation dictionaries such as:

        {"date": "...", "bidPrice": "..."}

    from accidentally being treated as fund records.
    """

    if not isinstance(
        value,
        dict,
    ):

        return False

    keys = {
        str(key).lower()
        for key in value.keys()
    }

    identity_hits = (
        keys
        & FUND_IDENTITY_KEYS
    )

    if not identity_hits:

        return False

    # Observation objects are explicitly excluded.
    if (
        "date" in keys
        and (
            "bidprice" in keys
            or "price" in keys
        )
        and not (
            "excelrow" in keys
            or "excel_row" in keys
        )
    ):

        return False

    return True


def flatten_fund_records(
    value,
    *,
    inherited_identifier: str | None = None,
) -> list[dict]:
    """
    Safely normalize fund records from arbitrary nested JSON.

    This is specifically designed to avoid:

        AttributeError:
        'list' object has no attribute 'get'

    It recursively handles:

        list
        dict
        nested funds lists
        dictionaries keyed by fund identifier

    but only accepts dictionaries that look like actual
    fund records.
    """

    results: list[dict] = []

    if isinstance(
        value,
        list,
    ):

        for item in value:

            results.extend(
                flatten_fund_records(
                    item,
                    inherited_identifier=
                        inherited_identifier,
                )
            )

        return results

    if not isinstance(
        value,
        dict,
    ):

        return results

    if is_probable_fund_record(
        value
    ):

        record = dict(
            value
        )

        if (
            inherited_identifier
            and not clean_text(
                record.get(
                    "fundIdentifier"
                )
            )
        ):

            record[
                "fundIdentifier"
            ] = inherited_identifier

        results.append(
            record
        )

        return results

    # --------------------------------------------------------
    # Known container keys.
    # --------------------------------------------------------

    container_keys = [
        "funds",
        "successfulFunds",
        "successful",
        "failedFunds",
        "failed",
        "failures",
        "records",
        "results",
        "data",
    ]

    for key in container_keys:

        if key not in value:

            continue

        child = value[
            key
        ]

        results.extend(
            flatten_fund_records(
                child
            )
        )

    # --------------------------------------------------------
    # Dictionary keyed by fund identifier.
    #
    # Example:
    #
    # "funds": {
    #     "D39F": {...}
    # }
    #
    # This section is only reached when the current dict itself
    # was not already recognized as a fund record.
    # --------------------------------------------------------

    for key, child in value.items():

        if key in container_keys:

            continue

        if isinstance(
            child,
            (dict, list),
        ):

            child_results = (
                flatten_fund_records(
                    child,
                    inherited_identifier=
                        str(key),
                )
            )

            results.extend(
                child_results
            )

    return results


def deduplicate_fund_records(
    records: list[dict],
) -> list[dict]:
    """
    Deduplicate fund records deterministically.

    Primary identity:
        excelRow

    Secondary identity:
        fundIdentifier

    This only prevents duplicated metadata records.
    It does not deduplicate historical observations.
    """

    result: list[dict] = []

    seen_rows = set()

    seen_identifiers = set()

    for record in records:

        if not isinstance(
            record,
            dict,
        ):

            continue

        excel_row = record.get(
            "excelRow"
        )

        identifier = clean_text(
            record.get(
                "fundIdentifier"
            )
        )

        if excel_row is not None:

            try:

                row_key = int(
                    excel_row
                )

            except (
                ValueError,
                TypeError,
            ):

                row_key = None

        else:

            row_key = None

        if (
            row_key is not None
            and row_key in seen_rows
        ):

            continue

        if (
            identifier
            and identifier in seen_identifiers
        ):

            continue

        if row_key is not None:

            seen_rows.add(
                row_key
            )

        if identifier:

            seen_identifiers.add(
                identifier
            )

        result.append(
            record
        )

    return result


# ============================================================
# SUCCESSFUL FUND EXTRACTION
# ============================================================

def extract_successful_funds(
    run_summary,
    all_funds,
) -> list[dict]:
    """
    Extract successful fund records robustly.

    Priority:

        1. all_funds.json -> funds
        2. run_summary.json -> successfulFunds

    The consolidated all_funds file is preferred because it
    contains the complete successful fund records.
    """

    candidates: list[dict] = []

    # --------------------------------------------------------
    # Preferred source: all_funds.json
    # --------------------------------------------------------

    if isinstance(
        all_funds,
        dict,
    ):

        all_funds_container = (
            all_funds.get(
                "funds"
            )
        )

        candidates.extend(
            flatten_fund_records(
                all_funds_container
            )
        )

    # --------------------------------------------------------
    # Fallback/source cross-check: run_summary.
    # --------------------------------------------------------

    if isinstance(
        run_summary,
        dict,
    ):

        successful_container = (
            run_summary.get(
                "successfulFunds"
            )
        )

        candidates.extend(
            flatten_fund_records(
                successful_container
            )
        )

    candidates = (
        deduplicate_fund_records(
            candidates
        )
    )

    return candidates


# ============================================================
# FAILED FUND EXTRACTION
# ============================================================

def extract_failed_funds(
    run_summary,
    all_funds,
) -> list[dict]:

    candidates: list[dict] = []

    # --------------------------------------------------------
    # run_summary
    # --------------------------------------------------------

    if isinstance(
        run_summary,
        dict,
    ):

        for key in (
            "failedFunds",
            "failed",
            "failures",
        ):

            if key in run_summary:

                candidates.extend(
                    flatten_fund_records(
                        run_summary.get(
                            key
                        )
                    )
                )

    # --------------------------------------------------------
    # all_funds
    # --------------------------------------------------------

    if isinstance(
        all_funds,
        dict,
    ):

        for key in (
            "failedFunds",
            "failed",
            "failures",
        ):

            if key in all_funds:

                candidates.extend(
                    flatten_fund_records(
                        all_funds.get(
                            key
                        )
                    )
                )

    # --------------------------------------------------------
    # Failed records may not contain fundIdentifier,
    # but should contain excelRow.
    #
    # flatten_fund_records recognizes excelRow as identity.
    # --------------------------------------------------------

    result: list[dict] = []

    seen_rows = set()

    for record in candidates:

        if not isinstance(
            record,
            dict,
        ):

            continue

        row = record.get(
            "excelRow"
        )

        if row is None:

            continue

        try:

            row = int(
                row
            )

        except (
            ValueError,
            TypeError,
        ):

            continue

        if row in seen_rows:

            continue

        seen_rows.add(
            row
        )

        result.append(
            record
        )

    return result


# ============================================================
# FIND SUMMARY RECORD
# ============================================================

def get_summary_excel_row(
    record: dict,
) -> int | None:

    if not isinstance(
        record,
        dict,
    ):

        return None

    value = record.get(
        "excelRow"
    )

    if value is None:

        value = record.get(
            "excel_row"
        )

    if value is None:

        value = record.get(
            "row"
        )

    try:

        return int(
            value
        )

    except (
        ValueError,
        TypeError,
    ):

        return None


# ============================================================
# SUSPICIOUS DATA DETECTION
# ============================================================

SUSPICIOUS_TERMS = (
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
)


def find_suspicious_keys(
    value,
    path: str = "",
) -> list[str]:
    """
    Search JSON recursively for keys indicating fabricated,
    synthetic, estimated, simulated, interpolated, or
    carry-forward data.

    This is intentionally strict.
    """

    findings: list[str] = []

    if isinstance(
        value,
        dict,
    ):

        for key, child in value.items():

            key_text = clean_text(
                key
            ).lower()

            for term in SUSPICIOUS_TERMS:

                if term in key_text:

                    findings.append(
                        (
                            f"{path}.{key}"
                            if path
                            else str(key)
                        )
                    )

                    break

            child_path = (
                f"{path}.{key}"
                if path
                else str(key)
            )

            findings.extend(
                find_suspicious_keys(
                    child,
                    child_path,
                )
            )

    elif isinstance(
        value,
        list,
    ):

        for index, child in enumerate(
            value
        ):

            child_path = (
                f"{path}[{index}]"
            )

            findings.extend(
                find_suspicious_keys(
                    child,
                    child_path,
                )
            )

    return findings


# ============================================================
# INDIVIDUAL FUND VALIDATION
# ============================================================

def validate_fund(
    excel_fund: dict,
    successful_summary: dict,
) -> tuple[bool, list[str], dict]:
    """
    Validate one successful fund.

    Returns:

        passed
        errors
        details
    """

    errors: list[str] = []

    details: dict = {
        "excelRow":
            excel_fund.get(
                "excelRow"
            ),

        "checks": {},
    }

    # --------------------------------------------------------
    # Defensive type check.
    #
    # This is the specific protection against the previous
    # "'list' object has no attribute 'get'" failure.
    # --------------------------------------------------------

    if not isinstance(
        successful_summary,
        dict,
    ):

        errors.append(
            "Successful fund record is not an object."
        )

        return (
            False,
            errors,
            details,
        )

    excel_row = excel_fund.get(
        "excelRow"
    )

    fund_directory = (
        get_fund_directory(
            excel_row
        )
    )

    if fund_directory is None:

        errors.append(
            "Individual fund output directory "
            "was not found."
        )

        return (
            False,
            errors,
            details,
        )

    details[
        "fundDirectory"
    ] = str(
        fund_directory
    )

    # --------------------------------------------------------
    # Required individual files.
    # --------------------------------------------------------

    required_files = [
        "summary.json",
        "bid_history.json",
        "pagination.json",
        "prudential_fund.json",
    ]

    loaded = {}

    for filename in required_files:

        path = (
            fund_directory
            / filename
        )

        if not path.exists():

            errors.append(
                f"Missing required file: {filename}"
            )

            continue

        try:

            loaded[
                filename
            ] = load_json(
                path
            )

        except Exception as error:

            errors.append(
                f"Could not load {filename}: {error}"
            )

    if errors:

        return (
            False,
            errors,
            details,
        )

    summary = loaded[
        "summary.json"
    ]

    bid_history = loaded[
        "bid_history.json"
    ]

    pagination = loaded[
        "pagination.json"
    ]

    prudential = loaded[
        "prudential_fund.json"
    ]

    # --------------------------------------------------------
    # All top-level structures must be dictionaries where
    # expected.
    # --------------------------------------------------------

    if not isinstance(
        summary,
        dict,
    ):

        errors.append(
            "summary.json is not an object."
        )

    if not isinstance(
        bid_history,
        dict,
    ):

        errors.append(
            "bid_history.json is not an object."
        )

    if not isinstance(
        pagination,
        list,
    ):

        errors.append(
            "pagination.json must be a list."
        )

    if not isinstance(
        prudential,
        dict,
    ):

        errors.append(
            "prudential_fund.json is not an object."
        )

    if errors:

        return (
            False,
            errors,
            details,
        )

    # ========================================================
    # EXCEL IDENTITY
    # ========================================================

    expected_excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    actual_excel_row = (
        successful_summary.get(
            "excelRow"
        )
    )

    if actual_excel_row is not None:

        try:

            actual_excel_row = int(
                actual_excel_row
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                "Successful record has invalid excelRow."
            )

    if (
        actual_excel_row is not None
        and actual_excel_row
        != expected_excel_row
    ):

        errors.append(
            "Successful record excelRow does not "
            "match Excel universe."
        )

    # ========================================================
    # FUND IDENTITY
    # ========================================================

    expected_prudential_name = clean_text(
        successful_summary.get(
            "fundName"
        )
        or successful_summary.get(
            "prudentialFundName"
        )
        or prudential.get(
            "fundName"
        )
    )

    actual_prudential_name = clean_text(
        prudential.get(
            "fundName"
        )
    )

    if not actual_prudential_name:

        errors.append(
            "Prudential fund name is missing."
        )

    if (
        expected_prudential_name
        and normalize_text(
            expected_prudential_name
        )
        != normalize_text(
            actual_prudential_name
        )
    ):

        errors.append(
            "Prudential fund name mismatch."
        )

    expected_identifier = clean_text(
        successful_summary.get(
            "fundIdentifier"
        )
        or prudential.get(
            "fundIdentifier"
        )
    )

    actual_identifier = clean_text(
        prudential.get(
            "fundIdentifier"
        )
    )

    if not actual_identifier:

        errors.append(
            "Prudential fundIdentifier is missing."
        )

    elif (
        expected_identifier
        and actual_identifier
        != expected_identifier
    ):

        errors.append(
            "fundIdentifier mismatch."
        )

    expected_fund_code = clean_text(
        successful_summary.get(
            "fundCode"
        )
        or prudential.get(
            "fundCode"
        )
    )

    actual_fund_code = clean_text(
        prudential.get(
            "fundCode"
        )
    )

    if (
        expected_fund_code
        and actual_fund_code
        != expected_fund_code
    ):

        errors.append(
            "Fund code mismatch."
        )

    # ========================================================
    # PRUACCESS IDENTITY
    # ========================================================

    excel_pruaccess_name = clean_text(
        excel_fund.get(
            "pruAccessName"
        )
    )

    actual_pruaccess_name = clean_text(
        bid_history.get(
            "pruAccessFundName"
        )
    )

    if not excel_pruaccess_name:

        errors.append(
            "Excel PruAccess name is empty."
        )

    if not actual_pruaccess_name:

        errors.append(
            "Extracted PruAccess fund name is missing."
        )

    elif (
        normalize_text(
            actual_pruaccess_name
        )
        != normalize_text(
            excel_pruaccess_name
        )
    ):

        errors.append(
            "PruAccess fund name does not exactly "
            "match Excel Column B."
        )

    pruaccess_id = clean_text(
        bid_history.get(
            "fundId"
        )
    )

    if not pruaccess_id:

        errors.append(
            "PruAccess fund ID is missing."
        )

    # ========================================================
    # SOURCE / PRICE TYPE
    # ========================================================

    source = normalize_text(
        bid_history.get(
            "source"
        )
    )

    if source != "pruaccess":

        errors.append(
            "Historical source is not PruAccess."
        )

    price_type = normalize_text(
        bid_history.get(
            "priceType"
        )
    )

    if price_type != "bid":

        errors.append(
            "Historical price type is not BID."
        )

    # ========================================================
    # OBSERVATIONS
    # ========================================================

    observations = bid_history.get(
        "observations"
    )

    if not isinstance(
        observations,
        list,
    ):

        errors.append(
            "bid_history.observations is not a list."
        )

        observations = []

    observation_count = len(
        observations
    )

    stored_observation_count = (
        bid_history.get(
            "observationCount"
        )
    )

    if stored_observation_count is not None:

        try:

            stored_observation_count = int(
                stored_observation_count
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                "bid_history observationCount is invalid."
            )

        else:

            if (
                stored_observation_count
                != observation_count
            ):

                errors.append(
                    "bid_history observationCount "
                    "does not match actual observations."
                )

    # --------------------------------------------------------
    # Successful summary count.
    # --------------------------------------------------------

    summary_count = (
        successful_summary.get(
            "observationCount"
        )
        or successful_summary.get(
            "historicalObservationCount"
        )
    )

    if summary_count is not None:

        try:

            summary_count = int(
                summary_count
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                "Successful summary observation count "
                "is invalid."
            )

        else:

            if (
                summary_count
                != observation_count
            ):

                errors.append(
                    "Successful summary observation count "
                    "does not match actual observations."
                )

    # ========================================================
    # DATE / PRICE VALIDATION
    # ========================================================

    parsed_dates: list[datetime] = []

    raw_dates: list[str] = []

    duplicate_dates = set()

    duplicate_records = set()

    previous_date = None

    for index, observation in enumerate(
        observations,
        start=1,
    ):

        if not isinstance(
            observation,
            dict,
        ):

            errors.append(
                f"Observation {index} is not an object."
            )

            continue

        date_value = clean_text(
            observation.get(
                "date"
            )
        )

        bid_value = clean_text(
            observation.get(
                "bidPrice"
            )
        )

        if not date_value:

            errors.append(
                f"Observation {index} has no date."
            )

            continue

        if not bid_value:

            errors.append(
                f"Observation {index} has no BID price."
            )

            continue

        raw_dates.append(
            date_value
        )

        try:

            parsed_date = parse_date(
                date_value
            )

            parsed_dates.append(
                parsed_date
            )

        except ValueError as error:

            errors.append(
                f"Observation {index}: {error}"
            )

            continue

        try:

            price = parse_price(
                bid_value
            )

        except ValueError as error:

            errors.append(
                f"Observation {index}: {error}"
            )

            continue

        # ----------------------------------------------------
        # Strict descending order.
        # ----------------------------------------------------

        if (
            previous_date is not None
            and parsed_date
            >= previous_date
        ):

            errors.append(
                "Historical dates are not strictly "
                "descending at observation "
                f"{index}."
            )

        previous_date = parsed_date

        # ----------------------------------------------------
        # Duplicate date.
        # ----------------------------------------------------

        if date_value in raw_dates[:-1]:

            duplicate_dates.add(
                date_value
            )

        # ----------------------------------------------------
        # Duplicate date + price.
        # ----------------------------------------------------

        record_key = (
            date_value,
            bid_value,
        )

        if record_key in duplicate_records:

            pass

        duplicate_records.add(
            record_key
        )

    # --------------------------------------------------------
    # Duplicate date detection using normalized parsed dates.
    # --------------------------------------------------------

    seen_date_objects = set()

    for parsed_date in parsed_dates:

        if parsed_date in seen_date_objects:

            duplicate_dates.add(
                parsed_date.strftime(
                    "%d-%b-%Y"
                )
            )

        seen_date_objects.add(
            parsed_date
        )

    if duplicate_dates:

        errors.append(
            "Duplicate historical dates found: "
            + ", ".join(
                sorted(
                    str(value)
                    for value in duplicate_dates
                )
            )
        )

    # --------------------------------------------------------
    # Duplicate date/price records.
    # --------------------------------------------------------

    seen_records = set()

    duplicate_record_values = []

    for observation in observations:

        if not isinstance(
            observation,
            dict,
        ):

            continue

        key = (
            clean_text(
                observation.get(
                    "date"
                )
            ),
            clean_text(
                observation.get(
                    "bidPrice"
                )
            ),
        )

        if key in seen_records:

            duplicate_record_values.append(
                key
            )

        seen_records.add(
            key
        )

    if duplicate_record_values:

        errors.append(
            "Duplicate date/BID observations found."
        )

    # ========================================================
    # OBSERVATION RANGE
    # ========================================================

    inception_raw = clean_text(
        prudential.get(
            "inceptionDate"
        )
    )

    if not inception_raw:

        errors.append(
            "Prudential inception date is missing."
        )

    else:

        try:

            inception_date = parse_date(
                inception_raw
            )

        except ValueError as error:

            errors.append(
                f"Invalid Prudential inception date: "
                f"{error}"
            )

        else:

            if parsed_dates:

                oldest_date = min(
                    parsed_dates
                )

                newest_date = max(
                    parsed_dates
                )

                # ------------------------------------------------
                # Historical data must reach inception.
                #
                # PruAccess may not have a price exactly on the
                # inception date in every possible circumstance,
                # so allow only the first available observation
                # on the same date or earlier.
                # ------------------------------------------------

                if oldest_date > inception_date:

                    errors.append(
                        "Historical BID data does not "
                        "reach Prudential inception date. "
                        f"Oldest={oldest_date.date()}, "
                        f"Inception={inception_date.date()}."
                    )

                # Newest observation cannot be after the
                # requested end date.
                requested_end_raw = clean_text(
                    bid_history.get(
                        "endDate"
                    )
                )

                if requested_end_raw:

                    try:

                        requested_end = parse_date(
                            requested_end_raw
                        )

                    except ValueError as error:

                        errors.append(
                            f"Invalid PruAccess end date: "
                            f"{error}"
                        )

                    else:

                        if newest_date > requested_end:

                            errors.append(
                                "Historical newest observation "
                                "is after requested end date."
                            )

    # ========================================================
    # BID HISTORY DATES
    # ========================================================

    start_date_raw = clean_text(
        bid_history.get(
            "startDate"
        )
    )

    end_date_raw = clean_text(
        bid_history.get(
            "endDate"
        )
    )

    if inception_raw and start_date_raw:

        try:

            inception_date = parse_date(
                inception_raw
            )

            history_start = parse_date(
                start_date_raw
            )

            if inception_date != history_start:

                errors.append(
                    "PruAccess start date does not "
                    "match Prudential inception date."
                )

        except ValueError:

            pass

    # ========================================================
    # PAGINATION
    # ========================================================

    if not isinstance(
        pagination,
        list,
    ):

        pagination = []

    if not pagination:

        errors.append(
            "Pagination diagnostics are empty."
        )

    page_numbers = []

    page_row_total = 0

    for index, page_record in enumerate(
        pagination,
        start=1,
    ):

        if not isinstance(
            page_record,
            dict,
        ):

            errors.append(
                f"Pagination record {index} is not an object."
            )

            continue

        page_number = page_record.get(
            "page"
        )

        row_count = page_record.get(
            "rowCount"
        )

        try:

            page_number = int(
                page_number
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                f"Pagination record {index} "
                "has invalid page number."
            )

            continue

        try:

            row_count = int(
                row_count
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                f"Pagination page {page_number} "
                "has invalid row count."
            )

            continue

        page_numbers.append(
            page_number
        )

        page_row_total += (
            row_count
        )

        if page_record.get(
            "status"
        ) != "success":

            errors.append(
                f"Pagination page {page_number} "
                "is not marked success."
            )

        if row_count < 1:

            errors.append(
                f"Pagination page {page_number} "
                "has no rows."
            )

    # --------------------------------------------------------
    # Page sequence.
    # --------------------------------------------------------

    expected_page_numbers = list(
        range(
            1,
            len(page_numbers) + 1,
        )
    )

    if page_numbers != expected_page_numbers:

        errors.append(
            "Pagination page sequence is invalid. "
            f"Actual={page_numbers}, "
            f"Expected={expected_page_numbers}."
        )

    # --------------------------------------------------------
    # Full pages.
    #
    # Every page except final should contain 20 rows.
    # Final page should contain 1..20.
    # --------------------------------------------------------

    for index, page_record in enumerate(
        pagination
    ):

        if not isinstance(
            page_record,
            dict,
        ):

            continue

        page_number = page_record.get(
            "page"
        )

        row_count = page_record.get(
            "rowCount"
        )

        try:

            page_number = int(
                page_number
            )

            row_count = int(
                row_count
            )

        except (
            ValueError,
            TypeError,
        ):

            continue

        is_final = (
            index
            == len(pagination) - 1
        )

        if not is_final:

            if row_count != PAGE_SIZE:

                errors.append(
                    f"Pagination page {page_number} "
                    f"returned {row_count} rows; "
                    f"expected {PAGE_SIZE}."
                )

        else:

            if not (
                1
                <= row_count
                <= PAGE_SIZE
            ):

                errors.append(
                    f"Final pagination page "
                    f"{page_number} has invalid "
                    f"row count {row_count}."
                )

    # --------------------------------------------------------
    # Pagination total must equal observations.
    # --------------------------------------------------------

    if page_row_total != observation_count:

        errors.append(
            "Pagination row total does not match "
            "historical observation count. "
            f"Pagination={page_row_total}, "
            f"Observations={observation_count}."
        )

    # --------------------------------------------------------
    # Summary page count.
    # --------------------------------------------------------

    summary_pages = (
        successful_summary.get(
            "pagesExtracted"
        )
    )

    if summary_pages is None:

        summary_pages = summary.get(
            "pagesExtracted"
        )

    if summary_pages is not None:

        try:

            summary_pages = int(
                summary_pages
            )

            if (
                summary_pages
                != len(pagination)
            ):

                errors.append(
                    "Summary pagesExtracted does not "
                    "match pagination page count."
                )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                "pagesExtracted is invalid."
            )

    # ========================================================
    # CURRENT PRUDENTIAL BID VS LATEST PRUACCESS BID
    # ========================================================

    current_prudential_bid = clean_text(
        prudential.get(
            "bidPrice"
        )
    )

    latest_pruaccess_bid = ""

    if observations:

        latest_observation = observations[
            0
        ]

        if isinstance(
            latest_observation,
            dict,
        ):

            latest_pruaccess_bid = clean_text(
                latest_observation.get(
                    "bidPrice"
                )
            )

    if not current_prudential_bid:

        errors.append(
            "Current Prudential BID price is missing."
        )

    if not latest_pruaccess_bid:

        errors.append(
            "Latest PruAccess BID price is missing."
        )

    if (
        current_prudential_bid
        and latest_pruaccess_bid
    ):

        try:

            prudential_bid = parse_price(
                current_prudential_bid
            )

            pruaccess_bid = parse_price(
                latest_pruaccess_bid
            )

            difference = abs(
                prudential_bid
                - pruaccess_bid
            )

            details[
                "currentPrudentialBid"
            ] = prudential_bid

            details[
                "latestPruAccessBid"
            ] = pruaccess_bid

            details[
                "currentBidDifference"
            ] = difference

            if difference > PRICE_TOLERANCE:

                errors.append(
                    "Current Prudential BID and latest "
                    "PruAccess BID differ beyond tolerance. "
                    f"Prudential={prudential_bid}, "
                    f"PruAccess={pruaccess_bid}, "
                    f"Difference={difference}, "
                    f"Tolerance={PRICE_TOLERANCE}."
                )

        except ValueError as error:

            errors.append(
                f"Could not compare current BID prices: "
                f"{error}"
            )

    # ========================================================
    # SYNTHETIC / FABRICATED DATA CHECK
    # ========================================================

    suspicious_files = {
        "summary.json":
            summary,

        "bid_history.json":
            bid_history,

        "pagination.json":
            pagination,

        "prudential_fund.json":
            prudential,
    }

    suspicious_findings = []

    for filename, content in (
        suspicious_files.items()
    ):

        findings = find_suspicious_keys(
            content
        )

        for finding in findings:

            suspicious_findings.append(
                f"{filename}:{finding}"
            )

    if suspicious_findings:

        errors.append(
            "Suspicious generated/synthetic-style "
            "fields detected: "
            + ", ".join(
                suspicious_findings
            )
        )

    # ========================================================
    # FUND STATUS
    # ========================================================

    status = normalize_text(
        summary.get(
            "status"
        )
    )

    if status != "success":

        errors.append(
            "Individual fund summary status "
            "is not success."
        )

    # ========================================================
    # DETAILS
    # ========================================================

    details[
        "fundName"
    ] = actual_prudential_name

    details[
        "fundIdentifier"
    ] = actual_identifier

    details[
        "fundCode"
    ] = actual_fund_code

    details[
        "pruAccessFundId"
    ] = pruaccess_id

    details[
        "observationCount"
    ] = observation_count

    details[
        "pageCount"
    ] = len(
        pagination
    )

    details[
        "oldestObservation"
    ] = (
        observations[-1]
        if observations
        else None
    )

    details[
        "newestObservation"
    ] = (
        observations[0]
        if observations
        else None
    )

    details[
        "checks"
    ] = {
        "identity":
            not any(
                "mismatch" in error.lower()
                or "missing" in error.lower()
                for error in errors
            ),

        "source":
            source == "pruaccess",

        "priceType":
            price_type == "bid",

        "observations":
            observation_count > 0,

        "pagination":
            bool(
                pagination
            )
            and page_row_total
            == observation_count,

        "syntheticData":
            not bool(
                suspicious_findings
            ),
    }

    passed = (
        len(errors)
        == 0
    )

    return (
        passed,
        errors,
        details,
    )


# ============================================================
# CONSOLIDATED HISTORY VALIDATION
# ============================================================

def validate_consolidated_history(
    excel_funds: list[dict],
    all_bid_history,
) -> tuple[bool, list[str], dict]:
    """
    Validate all_bid_history.json.
    """

    errors: list[str] = []

    details = {
        "checks": {},
    }

    if not isinstance(
        all_bid_history,
        dict,
    ):

        return (
            False,
            [
                "all_bid_history.json is not an object."
            ],
            details,
        )

    source = normalize_text(
        all_bid_history.get(
            "source"
        )
    )

    if source != "pruaccess":

        errors.append(
            "Consolidated history source is not PruAccess."
        )

    price_type = normalize_text(
        all_bid_history.get(
            "priceType"
        )
    )

    if price_type != "bid":

        errors.append(
            "Consolidated history priceType is not BID."
        )

    funds_container = (
        all_bid_history.get(
            "funds"
        )
    )

    if not isinstance(
        funds_container,
        dict,
    ):

        errors.append(
            "all_bid_history.funds is not an object."
        )

        return (
            False,
            errors,
            details,
        )

    # --------------------------------------------------------
    # Expected successful fund count based on Excel cannot be
    # assumed because there may be extraction failures.
    # Compare against actual successful individual records
    # later.
    # --------------------------------------------------------

    consolidated_fund_count = len(
        funds_container
    )

    details[
        "consolidatedFundCount"
    ] = consolidated_fund_count

    total_observations = 0

    for identifier, record in (
        funds_container.items()
    ):

        if not isinstance(
            record,
            dict,
        ):

            errors.append(
                f"Consolidated fund {identifier} "
                "is not an object."
            )

            continue

        record_source = normalize_text(
            record.get(
                "source"
            )
        )

        record_price_type = normalize_text(
            record.get(
                "priceType"
            )
        )

        if record_source != "pruaccess":

            errors.append(
                f"Consolidated fund {identifier} "
                "source is not PruAccess."
            )

        if record_price_type != "bid":

            errors.append(
                f"Consolidated fund {identifier} "
                "priceType is not BID."
            )

        observations = record.get(
            "observations"
        )

        if not isinstance(
            observations,
            list,
        ):

            errors.append(
                f"Consolidated fund {identifier} "
                "observations is not a list."
            )

            continue

        observation_count = len(
            observations
        )

        stored_count = record.get(
            "observationCount"
        )

        if stored_count is not None:

            try:

                stored_count = int(
                    stored_count
                )

            except (
                ValueError,
                TypeError,
            ):

                errors.append(
                    f"Consolidated fund {identifier} "
                    "has invalid observationCount."
                )

            else:

                if (
                    stored_count
                    != observation_count
                ):

                    errors.append(
                        f"Consolidated fund {identifier} "
                        "observationCount mismatch."
                    )

        total_observations += (
            observation_count
        )

    details[
        "totalHistoricalBidObservations"
    ] = total_observations

    reported_total = (
        all_bid_history.get(
            "totalHistoricalBidObservations"
        )
    )

    if reported_total is not None:

        try:

            reported_total = int(
                reported_total
            )

        except (
            ValueError,
            TypeError,
        ):

            errors.append(
                "Consolidated totalHistoricalBidObservations "
                "is invalid."
            )

        else:

            if (
                reported_total
                != total_observations
            ):

                errors.append(
                    "Consolidated totalHistoricalBidObservations "
                    "does not match actual total."
                )

    passed = (
        len(errors)
        == 0
    )

    return (
        passed,
        errors,
        details,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "======================================================================"
    )

    print(
        "PruAccess All-Fund Historical BID Price Validation"
    )

    print(
        "======================================================================"
    )

    # ========================================================
    # LOAD MASTER DATA
    # ========================================================

    excel_funds = read_excel_funds()

    run_summary = load_json(
        RUN_SUMMARY_FILE
    )

    all_funds = load_json(
        ALL_FUNDS_FILE
    )

    all_bid_history = load_json(
        ALL_BID_HISTORY_FILE
    )

    excel_count = len(
        excel_funds
    )

    print(
        f"\nExcel fund universe: {excel_count}"
    )

    # ========================================================
    # BASIC RUN SUMMARY
    # ========================================================

    if not isinstance(
        run_summary,
        dict,
    ):

        raise RuntimeError(
            "run_summary.json is not an object."
        )

    reported_universe_count = (
        run_summary.get(
            "fundUniverseCount"
        )
    )

    if reported_universe_count is not None:

        try:

            reported_universe_count = int(
                reported_universe_count
            )

        except (
            ValueError,
            TypeError,
        ):

            reported_universe_count = None

            print(
                "WARNING: run_summary fundUniverseCount "
                "is not numeric."
            )

        if (
            reported_universe_count is not None
            and reported_universe_count
            != excel_count
        ):

            print(
                "ERROR: run_summary fundUniverseCount "
                "does not match Excel universe."
            )

    # ========================================================
    # NORMALIZE SUCCESSFUL FUNDS
    # ========================================================

    successful_funds = (
        extract_successful_funds(
            run_summary,
            all_funds,
        )
    )

    failed_funds = (
        extract_failed_funds(
            run_summary,
            all_funds,
        )
    )

    print(
        f"Normalized successful fund records: "
        f"{len(successful_funds)}"
    )

    print(
        f"Normalized failed fund records: "
        f"{len(failed_funds)}"
    )

    # ========================================================
    # INDEX SUCCESSFUL BY EXCEL ROW
    # ========================================================

    successful_by_row: dict[int, dict] = {}

    malformed_success_records = []

    for record in successful_funds:

        if not isinstance(
            record,
            dict,
        ):

            malformed_success_records.append(
                record
            )

            continue

        excel_row = get_summary_excel_row(
            record
        )

        if excel_row is None:

            # ------------------------------------------------
            # Do not crash. Record malformed structure.
            # ------------------------------------------------

            malformed_success_records.append(
                record
            )

            continue

        if excel_row in successful_by_row:

            print(
                f"WARNING: Duplicate successful "
                f"record for Excel row {excel_row}."
            )

        successful_by_row[
            excel_row
        ] = record

    if malformed_success_records:

        print(
            f"Malformed successful records ignored: "
            f"{len(malformed_success_records)}"
        )

    # ========================================================
    # INDEX FAILED BY EXCEL ROW
    # ========================================================

    failed_by_row: dict[int, dict] = {}

    for record in failed_funds:

        if not isinstance(
            record,
            dict,
        ):

            continue

        excel_row = get_summary_excel_row(
            record
        )

        if excel_row is None:

            continue

        failed_by_row[
            excel_row
        ] = record

    # ========================================================
    # COUNTS
    # ========================================================

    successful_count = len(
        successful_by_row
    )

    failed_count = len(
        failed_by_row
    )

    print(
        f"\nSuccessful funds: {successful_count}"
    )

    print(
        f"Failed funds: {failed_count}"
    )

    # ========================================================
    # COVERAGE CHECK
    # ========================================================

    excel_rows = {
        int(
            fund["excelRow"]
        )
        for fund in excel_funds
    }

    successful_rows = set(
        successful_by_row.keys()
    )

    failed_rows = set(
        failed_by_row.keys()
    )

    missing_rows = (
        excel_rows
        - successful_rows
        - failed_rows
    )

    unexpected_success_rows = (
        successful_rows
        - excel_rows
    )

    unexpected_failed_rows = (
        failed_rows
        - excel_rows
    )

    overall_errors: list[str] = []

    if missing_rows:

        overall_errors.append(
            "Excel rows are missing from both "
            "successful and failed extraction records: "
            + ", ".join(
                str(row)
                for row in sorted(
                    missing_rows
                )
            )
        )

    if unexpected_success_rows:

        overall_errors.append(
            "Successful records contain Excel rows "
            "not present in the master Excel universe: "
            + ", ".join(
                str(row)
                for row in sorted(
                    unexpected_success_rows
                )
            )
        )

    if unexpected_failed_rows:

        overall_errors.append(
            "Failed records contain Excel rows "
            "not present in the master Excel universe: "
            + ", ".join(
                str(row)
                for row in sorted(
                    unexpected_failed_rows
                )
            )
        )

    if (
        successful_count
        + failed_count
        != excel_count
    ):

        overall_errors.append(
            "Successful + failed fund counts do not "
            "cover the complete Excel universe."
        )

    # ========================================================
    # VALIDATE EACH EXCEL FUND
    # ========================================================

    validation_results = []

    passed_count = 0

    failed_validation_count = 0

    print(
        "\nValidating individual funds..."
    )

    for index, excel_fund in enumerate(
        excel_funds,
        start=1,
    ):

        excel_row = int(
            excel_fund[
                "excelRow"
            ]
        )

        print(
            f"\n[{index}/{excel_count}] "
            f"Excel row {excel_row}:"
        )

        # ----------------------------------------------------
        # Extraction failure.
        # ----------------------------------------------------

        if excel_row in failed_by_row:

            failure_record = failed_by_row[
                excel_row
            ]

            failure_error = clean_text(
                failure_record.get(
                    "error"
                )
            )

            errors = [
                "Fund extraction failed."
            ]

            if failure_error:

                errors.append(
                    f"Extractor error: {failure_error}"
                )

            validation_results.append(
                {
                    "excelRow":
                        excel_row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "pruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "status":
                        "failed",

                    "errors":
                        errors,
                }
            )

            failed_validation_count += 1

            print(
                "  FAIL - extraction failed"
            )

            continue

        # ----------------------------------------------------
        # Missing success/failure record.
        # ----------------------------------------------------

        if excel_row not in successful_by_row:

            errors = [
                "No successful or failed extraction "
                "record exists for this Excel row."
            ]

            validation_results.append(
                {
                    "excelRow":
                        excel_row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "pruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "status":
                        "failed",

                    "errors":
                        errors,
                }
            )

            failed_validation_count += 1

            print(
                "  FAIL - missing extraction record"
            )

            continue

        # ----------------------------------------------------
        # Validate actual successful fund.
        # ----------------------------------------------------

        successful_summary = (
            successful_by_row[
                excel_row
            ]
        )

        # ----------------------------------------------------
        # Defensive guard.
        # ----------------------------------------------------

        if not isinstance(
            successful_summary,
            dict,
        ):

            errors = [
                "Successful extraction record is not "
                "a dictionary."
            ]

            validation_results.append(
                {
                    "excelRow":
                        excel_row,

                    "status":
                        "failed",

                    "errors":
                        errors,
                }
            )

            failed_validation_count += 1

            print(
                "  FAIL - malformed successful record"
            )

            continue

        passed, errors, details = (
            validate_fund(
                excel_fund,
                successful_summary,
            )
        )

        validation_results.append(
            {
                "excelRow":
                    excel_row,

                "prudentialUrl":
                    excel_fund[
                        "prudentialUrl"
                    ],

                "pruAccessName":
                    excel_fund[
                        "pruAccessName"
                    ],

                "status":
                    (
                        "passed"
                        if passed
                        else "failed"
                    ),

                "errors":
                    errors,

                "details":
                    details,
            }
        )

        if passed:

            passed_count += 1

            print(
                "  PASS"
            )

            print(
                f"  Fund: "
                f"{details.get('fundName')}"
            )

            print(
                f"  Identifier: "
                f"{details.get('fundIdentifier')}"
            )

            print(
                f"  Observations: "
                f"{details.get('observationCount')}"
            )

            print(
                f"  Pages: "
                f"{details.get('pageCount')}"
            )

        else:

            failed_validation_count += 1

            print(
                "  FAIL"
            )

            for error in errors:

                print(
                    f"    - {error}"
                )

    # ========================================================
    # CONSOLIDATED HISTORY
    # ========================================================

    print(
        "\nValidating consolidated BID history..."
    )

    (
        consolidated_passed,
        consolidated_errors,
        consolidated_details,
    ) = validate_consolidated_history(
        excel_funds,
        all_bid_history,
    )

    if consolidated_passed:

        print(
            "  Consolidated history: PASS"
        )

    else:

        print(
            "  Consolidated history: FAIL"
        )

        for error in consolidated_errors:

            print(
                f"    - {error}"
            )

    # ========================================================
    # ALL FUNDS CONSOLIDATED FILE
    # ========================================================

    all_funds_errors: list[str] = []

    if not isinstance(
        all_funds,
        dict,
    ):

        all_funds_errors.append(
            "all_funds.json is not an object."
        )

    else:

        all_funds_funds = (
            all_funds.get(
                "funds"
            )
        )

        if not isinstance(
            all_funds_funds,
            list,
        ):

            all_funds_errors.append(
                "all_funds.json funds field "
                "is not a list."
            )

        else:

            all_funds_success_count = len(
                [
                    item
                    for item in all_funds_funds
                    if (
                        isinstance(
                            item,
                            dict,
                        )
                        and normalize_text(
                            item.get(
                                "status"
                            )
                        )
                        == "success"
                    )
                ]
            )

            if (
                all_funds_success_count
                != successful_count
            ):

                all_funds_errors.append(
                    "all_funds.json successful fund "
                    "count does not match normalized "
                    "successful extraction count."
                )

        all_funds_universe_count = (
            all_funds.get(
                "fundUniverseCount"
            )
        )

        if all_funds_universe_count is not None:

            try:

                all_funds_universe_count = int(
                    all_funds_universe_count
                )

            except (
                ValueError,
                TypeError,
            ):

                all_funds_errors.append(
                    "all_funds fundUniverseCount "
                    "is invalid."
                )

            else:

                if (
                    all_funds_universe_count
                    != excel_count
                ):

                    all_funds_errors.append(
                        "all_funds fundUniverseCount "
                        "does not match Excel universe."
                    )

    if all_funds_errors:

        print(
            "\nall_funds.json: FAIL"
        )

        for error in all_funds_errors:

            print(
                f"  - {error}"
            )

        overall_errors.extend(
            all_funds_errors
        )

    else:

        print(
            "\nall_funds.json: PASS"
        )

    # ========================================================
    # RUN SUMMARY COUNTS
    # ========================================================

    run_summary_errors: list[str] = []

    reported_successful = (
        run_summary.get(
            "successfulFundCount"
        )
    )

    reported_failed = (
        run_summary.get(
            "failedFundCount"
        )
    )

    if reported_successful is not None:

        try:

            reported_successful = int(
                reported_successful
            )

            if (
                reported_successful
                != successful_count
            ):

                run_summary_errors.append(
                    "run_summary successfulFundCount "
                    "does not match normalized successful "
                    "fund count."
                )

        except (
            ValueError,
            TypeError,
        ):

            run_summary_errors.append(
                "run_summary successfulFundCount "
                "is invalid."
            )

    if reported_failed is not None:

        try:

            reported_failed = int(
                reported_failed
            )

            if (
                reported_failed
                != failed_count
            ):

                run_summary_errors.append(
                    "run_summary failedFundCount "
                    "does not match normalized failed "
                    "fund count."
                )

        except (
            ValueError,
            TypeError,
        ):

            run_summary_errors.append(
                "run_summary failedFundCount "
                "is invalid."
            )

    if run_summary_errors:

        print(
            "\nrun_summary.json: FAIL"
        )

        for error in run_summary_errors:

            print(
                f"  - {error}"
            )

        overall_errors.extend(
            run_summary_errors
        )

    else:

        print(
            "\nrun_summary.json: PASS"
        )

    # ========================================================
    # TOTAL OBSERVATION CHECK
    # ========================================================

    actual_individual_total = 0

    for result in validation_results:

        if result.get(
            "status"
        ) != "passed":

            continue

        details = result.get(
            "details"
        )

        if not isinstance(
            details,
            dict,
        ):

            continue

        observation_count = details.get(
            "observationCount"
        )

        if observation_count is None:

            continue

        try:

            actual_individual_total += int(
                observation_count
            )

        except (
            ValueError,
            TypeError,
        ):

            pass

    consolidated_total = (
        consolidated_details.get(
            "totalHistoricalBidObservations"
        )
    )

    if consolidated_total is not None:

        try:

            consolidated_total = int(
                consolidated_total
            )

            if (
                consolidated_total
                != actual_individual_total
            ):

                overall_errors.append(
                    "Consolidated total historical "
                    "BID observations does not match "
                    "sum of passed individual funds."
                )

        except (
            ValueError,
            TypeError,
        ):

            overall_errors.append(
                "Consolidated total observation count "
                "is invalid."
            )

    # ========================================================
    # FINAL STATUS
    # ========================================================

    overall_passed = (
        len(overall_errors)
        == 0
        and failed_validation_count
        == 0
        and passed_count
        == excel_count
        and consolidated_passed
    )

    validation_output = {
        "status":
            (
                "PASS"
                if overall_passed
                else "FAIL"
            ),

        "validatedAtUtc":
            datetime.utcnow().isoformat()
            + "Z",

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "excelFundUniverseCount":
            excel_count,

        "successfulFundCount":
            successful_count,

        "failedExtractionFundCount":
            failed_count,

        "passedValidationFundCount":
            passed_count,

        "failedValidationFundCount":
            failed_validation_count,

        "malformedSuccessfulRecordCount":
            len(
                malformed_success_records
            ),

        "missingExcelRows":
            sorted(
                missing_rows
            ),

        "unexpectedSuccessfulRows":
            sorted(
                unexpected_success_rows
            ),

        "unexpectedFailedRows":
            sorted(
                unexpected_failed_rows
            ),

        "consolidatedHistory":
            {
                "passed":
                    consolidated_passed,

                "errors":
                    consolidated_errors,

                "details":
                    consolidated_details,
            },

        "overallErrors":
            overall_errors,

        "fundResults":
            validation_results,
    }

    save_json(
        VALIDATION_FILE,
        validation_output,
    )

    # ========================================================
    # FINAL REPORT
    # ========================================================

    print(
        "\n\n"
        "======================================================================"
    )

    print(
        "VALIDATION COMPLETE"
    )

    print(
        "======================================================================"
    )

    print(
        f"Excel fund universe: "
        f"{excel_count}"
    )

    print(
        f"Normalized successful funds: "
        f"{successful_count}"
    )

    print(
        f"Normalized failed extractions: "
        f"{failed_count}"
    )

    print(
        f"Passed individual validations: "
        f"{passed_count}"
    )

    print(
        f"Failed individual validations: "
        f"{failed_validation_count}"
    )

    print(
        f"Consolidated history: "
        f"{'PASS' if consolidated_passed else 'FAIL'}"
    )

    print(
        f"Overall validation: "
        f"{'PASS' if overall_passed else 'FAIL'}"
    )

    if overall_errors:

        print(
            "\nOverall errors:"
        )

        for error in overall_errors:

            print(
                f" - {error}"
            )

    print(
        f"\nValidation file:"
    )

    print(
        f" - {VALIDATION_FILE}"
    )

    print(
        "\n======================================================================"
    )

    if not overall_passed:

        raise SystemExit(
            1
        )


if __name__ == "__main__":
    main()
