#!/usr/bin/env python3

"""
VGrat FMS production fund-data builder.

PRODUCTION WORKFLOW
===================

1. Read the Excel master universe from Funds Links.xlsm.
2. Use test_pruaccess.py's already-tested Prudential extraction.
3. Use test_pruaccess.py's already-tested PruAccess extraction.
4. Process every Excel-master fund independently.
5. A successful fund immediately replaces its previous production data.
6. A failed fund retains its previous valid production data.
7. Failed funds are flagged in the production run metadata.
8. Every production fund contains a freshness/status flag.
9. Mixed current and retained data is allowed.
10. No email system is used.

FRESHNESS RULE
==============

SUCCESSFUL CURRENT-RUN FUND:

    productionStatus:
        "updated"

    dataFreshness.status:
        "current"

    dataFreshness.isStale:
        false

    dataFreshness.lastSuccessfulUpdateUtc:
        current successful extraction timestamp


FAILED FUND WITH PREVIOUS DATA:

    productionStatus:
        "retained_previous"

    dataFreshness.status:
        "stale"

    dataFreshness.isStale:
        true

    dataFreshness.lastSuccessfulUpdateUtc:
        previous successful extraction timestamp


FAILED FUND WITH NO PREVIOUS DATA:

    productionStatus:
        "unavailable"

    dataFreshness.status:
        "unavailable"

    dataFreshness.isStale:
        true

    dataFreshness.lastSuccessfulUpdateUtc:
        null


IMPORTANT
=========

Freshness is a metadata flag only.

The actual historical BID data is NEVER modified to make it
appear current.

No missing dates are fabricated.

No prices are interpolated.

No prices are estimated.

No carry-forward values are written into raw historical data.

The actual extraction and validation rules remain in:

    test_pruaccess.py

This production script imports those tested functions rather than
rewriting them.

EXIT STATUS
===========

0:
    All Excel-master funds passed.

1:
    One or more funds failed.

A failed fund does NOT stop the processing of the remaining funds.
"""


from __future__ import annotations

import asyncio
import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from test_pruaccess import (
    EXCEL_FILE,
    OUTPUT_DIR,
    BROWSER_HEADLESS,
    BrowserContext,
    Browser,
    async_playwright,
    clean_text,
    get_fund_options,
    extract_single_fund,
    read_excel_funds,
    save_json,
    utc_now_iso,
)


# ============================================================
# PRODUCTION DIRECTORIES
# ============================================================

PRODUCTION_DIR = (
    OUTPUT_DIR
    / "production"
)

PRODUCTION_FUNDS_DIR = (
    PRODUCTION_DIR
    / "funds"
)

RUNS_DIR = (
    OUTPUT_DIR
    / "runs"
)

LEGACY_FUNDS_DIR = (
    OUTPUT_DIR
    / "funds"
)


# ============================================================
# FRESHNESS STATUS VALUES
# ============================================================

FRESHNESS_CURRENT = "current"
FRESHNESS_STALE = "stale"
FRESHNESS_UNAVAILABLE = "unavailable"


# ============================================================
# BASIC FILE HELPERS
# ============================================================

def remove_directory(
    path: Path,
) -> None:

    if path.exists():

        shutil.rmtree(
            path
        )


def copy_directory(
    source: Path,
    destination: Path,
) -> None:

    if not source.exists():

        raise FileNotFoundError(
            f"Source directory does not exist: {source}"
        )

    destination.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if destination.exists():

        remove_directory(
            destination
        )

    shutil.copytree(
        source,
        destination,
    )


def load_json_file(
    path: Path,
) -> Any:

    return json.loads(
        path.read_text(
            encoding="utf-8"
        )
    )


def find_fund_directory_by_excel_row(
    base_directory: Path,
    excel_row: int,
) -> Path | None:

    if not base_directory.exists():

        return None

    prefix = (
        f"{excel_row}_"
    )

    matches = [
        path
        for path in base_directory.iterdir()
        if path.is_dir()
        and path.name.startswith(prefix)
        and not path.name.endswith("_failed")
    ]

    if not matches:

        return None

    matches.sort(
        key=lambda path:
            path.name
    )

    return matches[0]


# ============================================================
# FRESHNESS HELPERS
# ============================================================

def utc_now_iso_fallback() -> str:

    return (
        datetime.now(
            timezone.utc
        )
        .isoformat()
    )


def get_previous_successful_update_time(
    directory: Path,
) -> str | None:

    """
    Retrieve the previous successful extraction timestamp.

    Newer production folders contain:

        production_metadata.json

    Older folders may not contain that file.

    In that case, fall back to the summary generatedAtUtc value
    where available.
    """

    metadata_file = (
        directory
        / "production_metadata.json"
    )

    if metadata_file.exists():

        try:

            metadata = load_json_file(
                metadata_file
            )

            value = clean_text(
                metadata.get(
                    "dataFreshness",
                    {}
                ).get(
                    "lastSuccessfulUpdateUtc"
                )
            )

            if value:

                return value

        except Exception:

            pass

    summary_file = (
        directory
        / "summary.json"
    )

    if summary_file.exists():

        try:

            summary = load_json_file(
                summary_file
            )

            value = clean_text(
                summary.get(
                    "lastSuccessfulUpdateUtc"
                )
            )

            if value:

                return value

            value = clean_text(
                summary.get(
                    "generatedAtUtc"
                )
            )

            if value:

                return value

        except Exception:

            pass

    return None


def build_current_freshness_metadata(
    successful_update_utc: str,
) -> dict[str, Any]:

    return {
        "status":
            FRESHNESS_CURRENT,

        "isStale":
            False,

        "lastSuccessfulUpdateUtc":
            successful_update_utc,

        "currentRunUpdated":
            True,

        "reason":
            "Fund data was successfully extracted and validated "
            "during the current production run.",
    }


def build_stale_freshness_metadata(
    last_successful_update_utc: str | None,
    failure_utc: str,
) -> dict[str, Any]:

    return {
        "status":
            FRESHNESS_STALE,

        "isStale":
            True,

        "lastSuccessfulUpdateUtc":
            last_successful_update_utc,

        "currentRunUpdated":
            False,

        "staleSinceRunUtc":
            failure_utc,

        "reason":
            "The fund failed the current production run and "
            "its previous valid production data was retained.",
    }


def build_unavailable_freshness_metadata(
    failure_utc: str,
) -> dict[str, Any]:

    return {
        "status":
            FRESHNESS_UNAVAILABLE,

        "isStale":
            True,

        "lastSuccessfulUpdateUtc":
            None,

        "currentRunUpdated":
            False,

        "staleSinceRunUtc":
            failure_utc,

        "reason":
            "The fund has no previous valid production data "
            "available.",
    }


# ============================================================
# PRODUCTION FUND DIRECTORY
# ============================================================

def production_fund_directory(
    result: dict[str, Any],
) -> Path:

    prudential = result[
        "prudential"
    ]

    excel_row = result[
        "excelRow"
    ]

    identifier = clean_text(
        prudential.get(
            "fundIdentifier"
        )
    )

    fund_code = clean_text(
        prudential.get(
            "fundCode"
        )
    )

    fund_name = clean_text(
        prudential.get(
            "fundName"
        )
    )

    safe_name = (
        identifier
        or fund_code
        or fund_name
        or "unknown"
    )

    safe_name = (
        safe_name
        .replace("/", "_")
        .replace("\\", "_")
        .replace(":", "_")
        .replace("*", "_")
        .replace("?", "_")
        .replace('"', "_")
        .replace("<", "_")
        .replace(">", "_")
        .replace("|", "_")
    )

    return (
        PRODUCTION_FUNDS_DIR
        / f"{excel_row}_{safe_name}"
    )


# ============================================================
# SAVE COMPLETE FUND RESULT
# ============================================================

def save_fund_result_directory(
    result: dict[str, Any],
    directory: Path,
) -> None:

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    prudential = result[
        "prudential"
    ]

    save_json(
        directory
        / "prudential_fund.json",
        prudential,
    )

    save_json(
        directory
        / "date_windows.json",
        result[
            "dateWindows"
        ],
    )

    save_json(
        directory
        / "window_diagnostics.json",
        result[
            "windowDiagnostics"
        ],
    )

    save_json(
        directory
        / "inception_validation.json",
        result[
            "inceptionValidation"
        ],
    )

    save_json(
        directory
        / "window_results.json",
        result[
            "windowResults"
        ],
    )

    save_json(
        directory
        / "bid_history.json",
        result[
            "bidHistory"
        ],
    )

    save_json(
        directory
        / "summary.json",
        result[
            "summary"
        ],
    )

    save_json(
        directory
        / "production_metadata.json",
        {
            "productionStatus":
                result.get(
                    "productionStatus"
                ),

            "dataFreshness":
                result.get(
                    "dataFreshness"
                ),
        },
    )


# ============================================================
# LOAD PREVIOUS PRODUCTION FUND
# ============================================================

def load_previous_fund(
    directory: Path,
    excel_fund: dict[str, Any],
) -> dict[str, Any] | None:

    if not directory.exists():

        return None

    required_files = [
        "prudential_fund.json",
        "date_windows.json",
        "window_diagnostics.json",
        "inception_validation.json",
        "window_results.json",
        "bid_history.json",
        "summary.json",
    ]

    for filename in required_files:

        if not (
            directory
            / filename
        ).exists():

            return None

    prudential = load_json_file(
        directory
        / "prudential_fund.json"
    )

    date_windows = load_json_file(
        directory
        / "date_windows.json"
    )

    window_diagnostics = load_json_file(
        directory
        / "window_diagnostics.json"
    )

    inception_validation = load_json_file(
        directory
        / "inception_validation.json"
    )

    window_results = load_json_file(
        directory
        / "window_results.json"
    )

    bid_history = load_json_file(
        directory
        / "bid_history.json"
    )

    summary = load_json_file(
        directory
        / "summary.json"
    )

    last_successful_update = (
        get_previous_successful_update_time(
            directory
        )
    )

    pruaccess = {
        "fundName":
            summary.get(
                "matchedPruAccessName"
            ),

        "fundId":
            summary.get(
                "pruAccessFundId"
            ),

        "startDate":
            summary.get(
                "startDate"
            ),

        "endDate":
            summary.get(
                "endDate"
            ),

        "priceType":
            "BID",

        "pageSize":
            summary.get(
                "pageSize"
            ),

        "windowCount":
            summary.get(
                "windowCount"
            ),

        "observationCount":
            summary.get(
                "historicalObservationCount"
            ),
    }

    freshness = build_stale_freshness_metadata(
        last_successful_update,
        utc_now_iso(),
    )

    return {
        "status":
            "retained_previous",

        "excelRow":
            excel_fund[
                "excelRow"
            ],

        "prudentialUrl":
            excel_fund[
                "prudentialUrl"
            ],

        "excelPruAccessName":
            excel_fund[
                "pruAccessName"
            ],

        "prudential":
            prudential,

        "pruAccess":
            pruaccess,

        "dateWindows":
            date_windows,

        "windowResults":
            window_results,

        "windowDiagnostics":
            window_diagnostics,

        "inceptionValidation":
            inception_validation,

        "summary":
            summary,

        "bidHistory":
            bid_history,

        "productionStatus":
            "retained_previous",

        "dataFreshness":
            freshness,
    }


# ============================================================
# PUBLISH ONE SUCCESSFUL FUND
# ============================================================

def publish_successful_fund(
    result: dict[str, Any],
    run_staging_directory: Path,
) -> Path:

    target_directory = (
        production_fund_directory(
            result
        )
    )

    target_directory.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    staging_directory = (
        run_staging_directory
        / target_directory.name
    )

    if staging_directory.exists():

        remove_directory(
            staging_directory
        )

    save_fund_result_directory(
        result,
        staging_directory,
    )

    row_prefix = (
        f"{result['excelRow']}_"
    )

    old_directories = [
        path
        for path in PRODUCTION_FUNDS_DIR.glob(
            f"{row_prefix}*"
        )
        if path.is_dir()
    ]

    backup_directories = []

    try:

        for old_directory in old_directories:

            backup_directory = (
                run_staging_directory
                / "_backups"
                / old_directory.name
            )

            backup_directory.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            old_directory.rename(
                backup_directory
            )

            backup_directories.append(
                (
                    old_directory,
                    backup_directory,
                )
            )

        staging_directory.rename(
            target_directory
        )

    except Exception:

        if target_directory.exists():

            remove_directory(
                target_directory
            )

        for (
            old_directory,
            backup_directory,
        ) in backup_directories:

            if backup_directory.exists():

                old_directory.parent.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                backup_directory.rename(
                    old_directory
                )

        raise

    for (
        _old_directory,
        backup_directory,
    ) in backup_directories:

        if backup_directory.exists():

            remove_directory(
                backup_directory
            )

    return target_directory


# ============================================================
# INITIAL LEGACY PRODUCTION MIGRATION
# ============================================================

def seed_production_from_legacy(
    excel_funds: list[dict[str, Any]],
) -> dict[int, str]:

    """
    Seed existing validated test data into production.

    This is only used when production data does not already
    exist for an Excel-master fund.

    Legacy data is treated as retained/previous data until that
    fund successfully passes the current production run.
    """

    seeded = {}

    if not LEGACY_FUNDS_DIR.exists():

        return seeded

    PRODUCTION_FUNDS_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    for excel_fund in excel_funds:

        row = excel_fund[
            "excelRow"
        ]

        existing = (
            find_fund_directory_by_excel_row(
                PRODUCTION_FUNDS_DIR,
                row,
            )
        )

        if existing:

            continue

        legacy = (
            find_fund_directory_by_excel_row(
                LEGACY_FUNDS_DIR,
                row,
            )
        )

        if not legacy:

            continue

        destination = (
            PRODUCTION_FUNDS_DIR
            / legacy.name
        )

        copy_directory(
            legacy,
            destination,
        )

        seeded[
            row
        ] = legacy.name

    return seeded


# ============================================================
# BUILD FAILURE RECORD
# ============================================================

def build_failure_record(
    excel_fund: dict[str, Any],
    error_text: str,
    previous_directory: Path | None,
    failure_utc: str,
) -> dict[str, Any]:

    previous_data_available = (
        previous_directory
        is not None
    )

    fund_name = None
    fund_identifier = None
    fund_code = None
    last_successful_update_utc = None

    if previous_directory:

        try:

            previous_prudential = (
                load_json_file(
                    previous_directory
                    / "prudential_fund.json"
                )
            )

            fund_name = clean_text(
                previous_prudential.get(
                    "fundName"
                )
            )

            fund_identifier = clean_text(
                previous_prudential.get(
                    "fundIdentifier"
                )
            )

            fund_code = clean_text(
                previous_prudential.get(
                    "fundCode"
                )
            )

            last_successful_update_utc = (
                get_previous_successful_update_time(
                    previous_directory
                )
            )

        except Exception:

            pass

    if previous_data_available:

        freshness = (
            build_stale_freshness_metadata(
                last_successful_update_utc,
                failure_utc,
            )
        )

        production_status = (
            "retained_previous"
        )

    else:

        freshness = (
            build_unavailable_freshness_metadata(
                failure_utc
            )
        )

        production_status = (
            "unavailable"
        )

    return {
        "status":
            "failed",

        "excelRow":
            excel_fund[
                "excelRow"
            ],

        "prudentialUrl":
            excel_fund[
                "prudentialUrl"
            ],

        "excelPruAccessName":
            excel_fund[
                "pruAccessName"
            ],

        "fundName":
            fund_name,

        "fundIdentifier":
            fund_identifier,

        "fundCode":
            fund_code,

        "error":
            error_text,

        "previousProductionDataRetained":
            previous_data_available,

        "previousProductionDirectory":
            (
                str(
                    previous_directory
                )
                if previous_directory
                else None
            ),

        "productionStatus":
            production_status,

        "dataFreshness":
            freshness,

        "failedAtUtc":
            failure_utc,
    }


# ============================================================
# COLLECT CURRENT PRODUCTION DATASET
# ============================================================

def collect_current_production_funds(
    excel_funds: list[dict[str, Any]],
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
]:

    production_funds = []
    missing_or_invalid = []

    for excel_fund in excel_funds:

        row = excel_fund[
            "excelRow"
        ]

        directory = (
            find_fund_directory_by_excel_row(
                PRODUCTION_FUNDS_DIR,
                row,
            )
        )

        if not directory:

            missing_or_invalid.append(
                {
                    "excelRow":
                        row,

                    "status":
                        "unavailable",

                    "reason":
                        "No valid production fund data "
                        "is available.",
                }
            )

            continue

        try:

            production_result = (
                load_previous_fund(
                    directory,
                    excel_fund,
                )
            )

        except Exception as error:

            missing_or_invalid.append(
                {
                    "excelRow":
                        row,

                    "status":
                        "unavailable",

                    "reason":
                        str(error),

                    "directory":
                        str(directory),
                }
            )

            continue

        if production_result is None:

            missing_or_invalid.append(
                {
                    "excelRow":
                        row,

                    "status":
                        "unavailable",

                    "reason":
                        "Production fund directory is incomplete.",

                    "directory":
                        str(directory),
                }
            )

            continue

        # ----------------------------------------------------
        # IMPORTANT:
        #
        # load_previous_fund() returns retained_previous by
        # default. We must recover the actual freshness state
        # stored in production_metadata.json where available.
        # ----------------------------------------------------

        metadata_file = (
            directory
            / "production_metadata.json"
        )

        if metadata_file.exists():

            try:

                metadata = load_json_file(
                    metadata_file
                )

                stored_status = (
                    clean_text(
                        metadata.get(
                            "productionStatus"
                        )
                    )
                )

                stored_freshness = (
                    metadata.get(
                        "dataFreshness"
                    )
                )

                if stored_status:

                    production_result[
                        "productionStatus"
                    ] = stored_status

                if isinstance(
                    stored_freshness,
                    dict,
                ):

                    production_result[
                        "dataFreshness"
                    ] = stored_freshness

            except Exception:

                pass

        production_funds.append(
            production_result
        )

    return (
        production_funds,
        missing_or_invalid,
    )


# ============================================================
# MAIN
# ============================================================

async def main():

    print(
        "============================================================"
    )

    print(
        "VGRAT FMS PRODUCTION FUND DATA BUILDER"
    )

    print(
        "============================================================"
    )

    run_started = utc_now_iso()

    run_id = (
        datetime.utcnow()
        .strftime(
            "%Y%m%d_%H%M%S"
        )
        + "_"
        + uuid.uuid4().hex[:8]
    )

    run_directory = (
        RUNS_DIR
        / run_id
    )

    run_staging_directory = (
        run_directory
        / "staging"
    )

    run_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    run_staging_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # READ EXCEL MASTER
    # ========================================================

    funds = read_excel_funds()

    total_funds = len(
        funds
    )

    print()
    print(
        f"Production fund universe: "
        f"{total_funds}"
    )

    if total_funds == 0:

        raise RuntimeError(
            "Excel master fund universe is empty."
        )

    # ========================================================
    # ENSURE PRODUCTION DIRECTORIES
    # ========================================================

    PRODUCTION_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    PRODUCTION_FUNDS_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # SEED FIRST PRODUCTION RUN
    # ========================================================

    seeded = seed_production_from_legacy(
        funds
    )

    if seeded:

        print()
        print(
            "Existing validated test data seeded into "
            "production for:"
        )

        for row, directory_name in seeded.items():

            print(
                f" - Excel row {row}: "
                f"{directory_name}"
            )

    # ========================================================
    # RUN TRACKING
    # ========================================================

    successful_funds = []
    retained_funds = []
    failed_funds = []

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    async with async_playwright() as playwright:

        browser: Browser = (
            await playwright.chromium.launch(
                headless=BROWSER_HEADLESS
            )
        )

        context: BrowserContext = (
            await browser.new_context(
                viewport={
                    "width":
                        1440,

                    "height":
                        1000,
                }
            )
        )

        page = await context.new_page()

        # ====================================================
        # LOAD PRUACCESS OPTIONS
        # ====================================================

        print()
        print(
            "============================================================"
        )

        print(
            "LOADING PRUACCESS FUND OPTIONS"
        )

        print(
            "============================================================"
        )

        pruaccess_options = (
            await get_fund_options(
                page
            )
        )

        save_json(
            run_directory
            / "fund_options.json",
            {
                "source":
                    "PruAccess",

                "retrievedAtUtc":
                    utc_now_iso(),

                "count":
                    len(
                        pruaccess_options
                    ),

                "options":
                    pruaccess_options,
            },
        )

        # ====================================================
        # PROCESS EVERY FUND
        # ====================================================

        for fund_index, excel_fund in enumerate(
            funds,
            start=1,
        ):

            print()
            print(
                "############################################################"
            )

            print(
                f"FUND {fund_index} / {total_funds}"
            )

            print(
                f"EXCEL ROW "
                f"{excel_fund['excelRow']}"
            )

            print(
                "############################################################"
            )

            try:

                result = (
                    await extract_single_fund(
                        context,
                        page,
                        excel_fund,
                        pruaccess_options,
                    )
                )

                # ------------------------------------------------
                # SUCCESSFUL CURRENT-RUN DATA
                # ------------------------------------------------

                successful_update_utc = (
                    utc_now_iso()
                )

                result[
                    "productionStatus"
                ] = "updated"

                result[
                    "dataFreshness"
                ] = (
                    build_current_freshness_metadata(
                        successful_update_utc
                    )
                )

                # ------------------------------------------------
                # SAVE ONLY AFTER SUCCESSFUL VALIDATION
                # ------------------------------------------------

                published_directory = (
                    publish_successful_fund(
                        result,
                        run_staging_directory,
                    )
                )

                result[
                    "productionDirectory"
                ] = str(
                    published_directory
                )

                successful_funds.append(
                    result
                )

                print()
                print(
                    "============================================================"
                )

                print(
                    "PRODUCTION FUND UPDATED"
                )

                print(
                    "============================================================"
                )

                print(
                    f"Fund: "
                    f"{result['prudential']['fundName']}"
                )

                print(
                    f"Excel row: "
                    f"{result['excelRow']}"
                )

                print(
                    "Freshness: current"
                )

                print(
                    f"Last successful update: "
                    f"{successful_update_utc}"
                )

                print(
                    f"Production directory: "
                    f"{published_directory}"
                )

            except Exception as error:

                error_text = clean_text(
                    str(error)
                )

                failure_utc = utc_now_iso()

                # ------------------------------------------------
                # FIND PREVIOUS PRODUCTION DATA
                # ------------------------------------------------

                previous_directory = (
                    find_fund_directory_by_excel_row(
                        PRODUCTION_FUNDS_DIR,
                        excel_fund[
                            "excelRow"
                        ],
                    )
                )

                failure = build_failure_record(
                    excel_fund,
                    error_text,
                    previous_directory,
                    failure_utc,
                )

                failed_funds.append(
                    failure
                )

                # ------------------------------------------------
                # SAVE FAILURE DIAGNOSTICS
                # ------------------------------------------------

                failure_directory = (
                    run_directory
                    / "failed"
                    / f"{excel_fund['excelRow']}_failed"
                )

                failure_directory.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                save_json(
                    failure_directory
                    / "failure.json",
                    failure,
                )

                if (
                    failure[
                        "previousProductionDataRetained"
                    ]
                ):

                    retained_funds.append(
                        {
                            "excelRow":
                                excel_fund[
                                    "excelRow"
                                ],

                            "fundName":
                                failure[
                                    "fundName"
                                ],

                            "fundIdentifier":
                                failure[
                                    "fundIdentifier"
                                ],

                            "fundCode":
                                failure[
                                    "fundCode"
                                ],

                            "status":
                                "retained_previous",

                            "productionStatus":
                                "retained_previous",

                            "dataFreshness":
                                failure[
                                    "dataFreshness"
                                ],

                            "directory":
                                failure[
                                    "previousProductionDirectory"
                                ],

                            "reason":
                                error_text,
                        }
                    )

                print()
                print(
                    "============================================================"
                )

                if (
                    failure[
                        "previousProductionDataRetained"
                    ]
                ):

                    print(
                        "FUND FAILED - PREVIOUS DATA RETAINED"
                    )

                else:

                    print(
                        "FUND FAILED - NO PREVIOUS DATA"
                    )

                print(
                    "============================================================"
                )

                print(
                    f"Excel row: "
                    f"{excel_fund['excelRow']}"
                )

                print(
                    f"PruAccess name: "
                    f"{excel_fund['pruAccessName']}"
                )

                print(
                    f"Reason: "
                    f"{error_text}"
                )

                print(
                    f"Production status: "
                    f"{failure['productionStatus']}"
                )

                print(
                    f"Freshness: "
                    f"{failure['dataFreshness']['status']}"
                )

                print(
                    f"Last successful update: "
                    f"{failure['dataFreshness']['lastSuccessfulUpdateUtc']}"
                )

                print(
                    f"Previous data retained: "
                    f"{failure['previousProductionDataRetained']}"
                )

                # ------------------------------------------------
                # CONTINUE TO NEXT FUND
                # ------------------------------------------------

                continue

        await browser.close()

    # ========================================================
    # LOAD FINAL PRODUCTION DATASET
    # ========================================================

    production_funds = []
    unavailable_production_funds = []

    for excel_fund in funds:

        row = excel_fund[
            "excelRow"
        ]

        directory = (
            find_fund_directory_by_excel_row(
                PRODUCTION_FUNDS_DIR,
                row,
            )
        )

        if not directory:

            unavailable_production_funds.append(
                {
                    "excelRow":
                        row,

                    "fundName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "productionStatus":
                        "unavailable",

                    "dataFreshness":
                        build_unavailable_freshness_metadata(
                            utc_now_iso()
                        ),
                }
            )

            continue

        try:

            production_result = (
                load_previous_fund(
                    directory,
                    excel_fund,
                )
            )

            if production_result is None:

                unavailable_production_funds.append(
                    {
                        "excelRow":
                            row,

                        "fundName":
                            excel_fund[
                                "pruAccessName"
                            ],

                        "productionStatus":
                            "unavailable",

                        "dataFreshness":
                            build_unavailable_freshness_metadata(
                                utc_now_iso()
                            ),

                        "directory":
                            str(directory),
                    }
                )

                continue

            # ------------------------------------------------
            # Recover stored production metadata.
            # ------------------------------------------------

            metadata_file = (
                directory
                / "production_metadata.json"
            )

            if metadata_file.exists():

                try:

                    metadata = load_json_file(
                        metadata_file
                    )

                    if metadata.get(
                        "productionStatus"
                    ):

                        production_result[
                            "productionStatus"
                        ] = metadata[
                            "productionStatus"
                        ]

                    if isinstance(
                        metadata.get(
                            "dataFreshness"
                        ),
                        dict,
                    ):

                        production_result[
                            "dataFreshness"
                        ] = metadata[
                            "dataFreshness"
                        ]

                except Exception:

                    pass

            production_funds.append(
                production_result
            )

        except Exception as error:

            print(
                f"WARNING: Could not load production fund "
                f"for Excel row {row}: {error}"
            )

            unavailable_production_funds.append(
                {
                    "excelRow":
                        row,

                    "fundName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "productionStatus":
                        "unavailable",

                    "dataFreshness":
                        build_unavailable_freshness_metadata(
                            utc_now_iso()
                        ),

                    "error":
                        clean_text(
                            str(error)
                        ),

                    "directory":
                        str(directory),
                }
            )

    # ========================================================
    # BID HISTORY TOTALS
    # ========================================================

    production_bid_history = {}

    total_observations = 0
    total_windows = 0
    total_pages = 0

    current_fresh_count = 0
    stale_fresh_count = 0
    unavailable_count = len(
        unavailable_production_funds
    )

    updated_row_set = {
        item[
            "excelRow"
        ]
        for item
        in successful_funds
    }

    for fund in production_funds:

        prudential = fund[
            "prudential"
        ]

        identifier = clean_text(
            prudential.get(
                "fundIdentifier"
            )
        )

        fund_code = clean_text(
            prudential.get(
                "fundCode"
            )
        )

        key = (
            identifier
            or fund_code
            or str(
                fund[
                    "excelRow"
                ]
            )
        )

        production_bid_history[
            key
        ] = fund[
            "bidHistory"
        ]

        total_observations += int(
            fund[
                "pruAccess"
            ].get(
                "observationCount"
            )
            or 0
        )

        total_windows += int(
            fund[
                "pruAccess"
            ].get(
                "windowCount"
            )
            or 0
        )

        total_pages += int(
            fund[
                "summary"
            ].get(
                "pagesExtracted"
            )
            or 0
        )

        freshness_status = (
            fund.get(
                "dataFreshness",
                {}
            ).get(
                "status"
            )
        )

        if freshness_status == FRESHNESS_CURRENT:

            current_fresh_count += 1

        elif freshness_status == FRESHNESS_STALE:

            stale_fresh_count += 1

    # ========================================================
    # FINAL RUN STATUS
    # ========================================================

    run_finished = utc_now_iso()

    if failed_funds:

        run_status = (
            "partial_success"
        )

    else:

        run_status = (
            "success"
        )

    # ========================================================
    # FRESHNESS SUMMARY
    # ========================================================

    freshness_summary = {
        "currentFundCount":
            current_fresh_count,

        "staleFundCount":
            stale_fresh_count,

        "unavailableFundCount":
            unavailable_count,

        "currentDefinition":
            "Fund successfully extracted and validated during "
            "the current production run.",

        "staleDefinition":
            "Fund failed the current production run and previous "
            "valid production data was retained.",

        "unavailableDefinition":
            "Fund has no valid production data available.",

        "dayBasedStalenessThreshold":
            None,

        "note":
            "Freshness is currently based on successful current-run "
            "status rather than an arbitrary number-of-days threshold.",
    }

    # ========================================================
    # RUN SUMMARY
    # ========================================================

    run_summary = {
        "status":
            run_status,

        "runId":
            run_id,

        "startedAtUtc":
            run_started,

        "completedAtUtc":
            run_finished,

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "fundUniverseCount":
            total_funds,

        "successfulFundCount":
            len(
                successful_funds
            ),

        "retainedPreviousFundCount":
            len(
                retained_funds
            ),

        "failedFundCount":
            len(
                failed_funds
            ),

        "productionFundCount":
            len(
                production_funds
            ),

        "unavailableProductionFundCount":
            len(
                unavailable_production_funds
            ),

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "totalHistoricalBidObservations":
            total_observations,

        "freshness":
            freshness_summary,

        "publicationRule":
            {
                "successfulFundReplacesPrevious":
                    True,

                "failedFundRetainsPrevious":
                    True,

                "failedFundStopsEntireRun":
                    False,

                "failedFundContinuesProcessing":
                    True,

                "mixedOldAndNewDataAllowed":
                    True,
            },

        "freshnessRule":
            {
                "currentRunSuccess":
                    "current",

                "failedWithPreviousData":
                    "stale",

                "failedWithoutPreviousData":
                    "unavailable",

                "rawHistoricalDataModifiedForFreshness":
                    False,

                "syntheticDataAllowed":
                    False,

                "interpolationAllowed":
                    False,

                "estimationAllowed":
                    False,
            },

        "successfulFunds":
            [
                {
                    "excelRow":
                        fund[
                            "excelRow"
                        ],

                    "fundName":
                        fund[
                            "prudential"
                        ].get(
                            "fundName"
                        ),

                    "fundIdentifier":
                        fund[
                            "prudential"
                        ].get(
                            "fundIdentifier"
                        ),

                    "fundCode":
                        fund[
                            "prudential"
                        ].get(
                            "fundCode"
                        ),

                    "status":
                        "updated",

                    "productionStatus":
                        "updated",

                    "dataFreshness":
                        fund[
                            "dataFreshness"
                        ],
                }

                for fund
                in successful_funds
            ],

        "retainedFunds":
            retained_funds,

        "failedFunds":
            failed_funds,

        "unavailableProductionFunds":
            unavailable_production_funds,
    }

    save_json(
        PRODUCTION_DIR
        / "run_summary.json",
        run_summary,
    )

    save_json(
        run_directory
        / "run_summary.json",
        run_summary,
    )

    # ========================================================
    # PRODUCTION ALL FUNDS
    # ========================================================

    all_funds = {
        "source":
            "Funds Links.xlsm",

        "generatedAtUtc":
            run_finished,

        "runId":
            run_id,

        "status":
            run_status,

        "fundUniverseCount":
            total_funds,

        "productionFundCount":
            len(
                production_funds
            ),

        "updatedFundCount":
            len(
                successful_funds
            ),

        "retainedPreviousFundCount":
            len(
                retained_funds
            ),

        "failedFundCount":
            len(
                failed_funds
            ),

        "unavailableProductionFundCount":
            len(
                unavailable_production_funds
            ),

        "freshness":
            freshness_summary,

        "funds":
            production_funds,

        "failedFunds":
            failed_funds,

        "unavailableProductionFunds":
            unavailable_production_funds,
    }

    save_json(
        PRODUCTION_DIR
        / "all_funds.json",
        all_funds,
    )

    # ========================================================
    # PRODUCTION ALL BID HISTORY
    # ========================================================

    all_bid_history = {
        "source":
            "PruAccess",

        "priceType":
            "BID",

        "generatedAtUtc":
            run_finished,

        "runId":
            run_id,

        "fundCount":
            len(
                production_bid_history
            ),

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "totalHistoricalBidObservations":
            total_observations,

        "freshness":
            freshness_summary,

        "funds":
            production_bid_history,
    }

    save_json(
        PRODUCTION_DIR
        / "all_bid_history.json",
        all_bid_history,
    )

    # ========================================================
    # VALIDATION SUMMARY
    # ========================================================

    validation_summary = {
        "status":
            run_status,

        "generatedAtUtc":
            run_finished,

        "runId":
            run_id,

        "excelFundUniverse":
            total_funds,

        "updatedFunds":
            len(
                successful_funds
            ),

        "retainedPreviousFunds":
            len(
                retained_funds
            ),

        "failedFunds":
            len(
                failed_funds
            ),

        "productionFunds":
            len(
                production_funds
            ),

        "unavailableProductionFunds":
            len(
                unavailable_production_funds
            ),

        "totalHistoricalBidObservations":
            total_observations,

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "freshness":
            freshness_summary,

        "rules":
            {
                "automaticTenYearWindows":
                    True,

                "completePaginationValidation":
                    True,

                "noSyntheticData":
                    True,

                "noInterpolation":
                    True,

                "noEstimation":
                    True,

                "chronologicalReconstruction":
                    True,

                "fullHistoryToInceptionValidation":
                    True,

                "maximumInceptionDelayDays":
                    7,

                "dateGapsAllowed":
                    True,

                "duplicateDatesCauseFailure":
                    True,

                "exactPruAccessNameMatch":
                    True,

                "excelMasterUniverse":
                    True,

                "hardcodedFundCount":
                    False,

                "historicalPriceType":
                    "BID",

                "mixedOldAndNewProductionDataAllowed":
                    True,

                "freshnessTracking":
                    True,

                "freshnessBasedOnCurrentRunSuccess":
                    True,

                "dayBasedStalenessThreshold":
                    False,

                "rawDataChangedForFreshness":
                    False,
            },

        "failedFundsDetail":
            failed_funds,
    }

    save_json(
        PRODUCTION_DIR
        / "validation.json",
        validation_summary,
    )

    save_json(
        run_directory
        / "validation.json",
        validation_summary,
    )

    # ========================================================
    # FINAL CONSOLE SUMMARY
    # ========================================================

    print()
    print(
        "============================================================"
    )

    print(
        "PRODUCTION BUILD COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Excel fund universe: "
        f"{total_funds}"
    )

    print(
        f"New successful updates: "
        f"{len(successful_funds)}"
    )

    print(
        f"Current funds: "
        f"{current_fresh_count}"
    )

    print(
        f"Stale funds: "
        f"{stale_fresh_count}"
    )

    print(
        f"Unavailable funds: "
        f"{unavailable_count}"
    )

    print(
        f"Previous data retained: "
        f"{len(retained_funds)}"
    )

    print(
        f"Failed funds: "
        f"{len(failed_funds)}"
    )

    print(
        f"Production funds available: "
        f"{len(production_funds)}"
    )

    print(
        f"Total historical BID observations: "
        f"{total_observations}"
    )

    print(
        f"Production directory: "
        f"{PRODUCTION_DIR}"
    )

    print(
        f"Run diagnostics: "
        f"{run_directory}"
    )

    # ========================================================
    # FRESHNESS DETAILS
    # ========================================================

    print()
    print(
        "FRESHNESS"
    )

    print(
        "------------------------------------------------------------"
    )

    print(
        f"Current: "
        f"{current_fresh_count}"
    )

    print(
        f"Stale: "
        f"{stale_fresh_count}"
    )

    print(
        f"Unavailable: "
        f"{unavailable_count}"
    )

    # ========================================================
    # FAILURE DETAILS
    # ========================================================

    if failed_funds:

        print()
        print(
            "FAILED FUNDS"
        )

        print(
            "------------------------------------------------------------"
        )

        for failure in failed_funds:

            print(
                f"Row {failure['excelRow']}: "
                f"{failure.get('fundName') or failure['excelPruAccessName']}"
            )

            print(
                f"  Reason: "
                f"{failure['error']}"
            )

            print(
                f"  Production status: "
                f"{failure['productionStatus']}"
            )

            print(
                f"  Freshness: "
                f"{failure['dataFreshness']['status']}"
            )

            print(
                f"  Last successful update: "
                f"{failure['dataFreshness']['lastSuccessfulUpdateUtc']}"
            )

    # ========================================================
    # EXIT STATUS
    # ========================================================

    if failed_funds:

        print()
        print(
            "============================================================"
        )

        print(
            "PRODUCTION BUILD FINISHED WITH FAILED FUNDS"
        )

        print(
            "============================================================"
        )

        print(
            "Successful funds were updated."
        )

        print(
            "Failed funds were retained where previous data existed."
        )

        print(
            "Freshness flags were updated."
        )

        print(
            "No email notification was configured."
        )

        raise SystemExit(1)

    print()
    print(
        "============================================================"
    )

    print(
        "PRODUCTION BUILD SUCCESS"
    )

    print(
        "============================================================"
    )

    print(
        f"All {total_funds} Excel-master funds "
        "were successfully updated."
    )

    print(
        "All production funds are marked current."
    )

    print(
        "No email notification is configured."
    )

    print(
        "Done."
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
