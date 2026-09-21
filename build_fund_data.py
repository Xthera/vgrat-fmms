#!/usr/bin/env python3

"""
VGrat FMS production fund-data builder.

LOCATION
========

This file is intended to live in the repository root:

    build_fund_data.py

The tested extraction engine lives in:

    scripts/test_pruaccess.py

The Prudential capture/debug test lives in:

    scripts/test_prudential_fund.py


PRODUCTION WORKFLOW
===================

1. Read the Excel master universe from Funds Links.xlsm.
2. Import the already-tested extraction/validation engine from
   scripts/test_pruaccess.py.
3. Process every Excel-master fund independently.
4. A successful fund immediately replaces its previous production
   data for that fund.
5. A failed fund retains its previous valid production data.
6. Processing continues after a fund failure.
7. Successful and retained/stale funds may coexist in production.
8. Every production fund has explicit freshness metadata.
9. No email system is used.


FRESHNESS RULE
==============

CURRENT-RUN SUCCESS:

    productionStatus:
        "updated"

    dataFreshness.status:
        "current"

    dataFreshness.isStale:
        false

    dataFreshness.currentRunUpdated:
        true

    dataFreshness.lastSuccessfulUpdateUtc:
        current successful run timestamp


FAILED FUND WITH PREVIOUS VALID DATA:

    productionStatus:
        "retained_previous"

    dataFreshness.status:
        "stale"

    dataFreshness.isStale:
        true

    dataFreshness.currentRunUpdated:
        false

    dataFreshness.lastSuccessfulUpdateUtc:
        last successful extraction timestamp


FAILED FUND WITH NO PREVIOUS VALID DATA:

    productionStatus:
        "unavailable"

    dataFreshness.status:
        "unavailable"

    dataFreshness.isStale:
        true

    dataFreshness.currentRunUpdated:
        false

    dataFreshness.lastSuccessfulUpdateUtc:
        null


IMPORTANT DATA RULES
====================

Freshness is metadata only.

The actual historical BID data is NEVER modified merely to make
it appear fresh.

No missing dates are fabricated.

No prices are interpolated.

No prices are estimated.

No carry-forward values are written into raw historical data.

The extraction and validation rules remain in:

    scripts/test_pruaccess.py

This production script imports those tested functions rather than
rewriting the extraction engine.


EXIT STATUS
===========

0:
    All Excel-master funds passed.

1:
    One or more Excel-master funds failed.

A failed fund does NOT stop the remaining funds from processing.
"""


from __future__ import annotations

import asyncio
import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


# ============================================================
# IMPORT TESTED EXTRACTION ENGINE
# ============================================================

from scripts.test_pruaccess import (
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


# ============================================================
# FIND FUND DIRECTORY
# ============================================================

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
        and not path.name.startswith("_")
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

def get_previous_successful_update_time(
    directory: Path,
) -> str | None:

    """
    Read the last successful production update timestamp.

    New production data:

        production_metadata.json

    Older production/test data:

        summary.json

    Fallback priority:

        production_metadata.json
        summary.lastSuccessfulUpdateUtc
        summary.generatedAtUtc
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

            freshness = metadata.get(
                "dataFreshness",
                {}
            )

            if isinstance(
                freshness,
                dict,
            ):

                value = clean_text(
                    freshness.get(
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

        "staleSinceRunUtc":
            None,

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
            "The fund has no valid previous production data "
            "available.",
    }


def write_production_metadata(
    directory: Path,
    production_status: str,
    freshness: dict[str, Any],
) -> None:

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory
        / "production_metadata.json",
        {
            "productionStatus":
                production_status,

            "dataFreshness":
                freshness,
        },
    )


def mark_existing_fund_stale(
    directory: Path,
    failure_utc: str,
) -> dict[str, Any]:

    """
    Update ONLY production metadata.

    The actual Prudential data, PruAccess BID history, window
    results, and validation files are left untouched.
    """

    last_successful_update_utc = (
        get_previous_successful_update_time(
            directory
        )
    )

    freshness = (
        build_stale_freshness_metadata(
            last_successful_update_utc,
            failure_utc,
        )
    )

    write_production_metadata(
        directory,
        "retained_previous",
        freshness,
    )

    return freshness


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
# LOAD PRODUCTION METADATA
# ============================================================

def load_production_metadata(
    directory: Path,
) -> dict[str, Any] | None:

    metadata_file = (
        directory
        / "production_metadata.json"
    )

    if not metadata_file.exists():

        return None

    try:

        metadata = load_json_file(
            metadata_file
        )

    except Exception:

        return None

    if not isinstance(
        metadata,
        dict,
    ):

        return None

    return metadata


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

    metadata = (
        load_production_metadata(
            directory
        )
    )

    if metadata:

        production_status = (
            clean_text(
                metadata.get(
                    "productionStatus"
                )
            )
        )

        data_freshness = (
            metadata.get(
                "dataFreshness"
            )
        )

        if not isinstance(
            data_freshness,
            dict,
        ):

            data_freshness = None

    else:

        # Legacy/test data has no production metadata.
        # Treat it as previous/stale until successfully updated
        # by the current production run.

        production_status = (
            "retained_previous"
        )

        data_freshness = (
            build_stale_freshness_metadata(
                get_previous_successful_update_time(
                    directory
                ),
                utc_now_iso(),
            )
        )

    if not production_status:

        production_status = (
            "retained_previous"
        )

    if data_freshness is None:

        data_freshness = (
            build_stale_freshness_metadata(
                get_previous_successful_update_time(
                    directory
                ),
                utc_now_iso(),
            )
        )

    return {
        "status":
            production_status,

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
            production_status,

        "dataFreshness":
            data_freshness,
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
        and not path.name.startswith("_")
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
    Copy previously validated test output into the initial
    production dataset when no production data exists yet.

    Seeded legacy data is explicitly marked as previous/stale.
    A successful current production run will immediately replace
    it and mark it current.
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

        # ----------------------------------------------------
        # Seeded legacy data is NOT current production data.
        # Mark it stale immediately.
        # ----------------------------------------------------

        last_successful_update_utc = (
            get_previous_successful_update_time(
                destination
            )
        )

        freshness = (
            build_stale_freshness_metadata(
                last_successful_update_utc,
                utc_now_iso(),
            )
        )

        write_production_metadata(
            destination,
            "retained_previous",
            freshness,
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
    freshness: dict[str, Any],
) -> dict[str, Any]:

    previous_data_available = (
        previous_directory
        is not None
    )

    fund_name = None
    fund_identifier = None
    fund_code = None

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

        except Exception:

            pass

    if previous_data_available:

        production_status = (
            "retained_previous"
        )

    else:

        production_status = (
            "unavailable"
        )

        freshness = (
            build_unavailable_freshness_metadata(
                failure_utc
            )
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
        datetime.now(
            timezone.utc
        )
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
    # SEED INITIAL PRODUCTION DATA
    # ========================================================

    seeded = seed_production_from_legacy(
        funds
    )

    if seeded:

        print()
        print(
            "Existing validated test data seeded into "
            "production as previous/stale data:"
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
                # SUCCESSFUL CURRENT-RUN FUND
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
                # PUBLISH IMMEDIATELY
                #
                # Mixed current/previous data is intentionally
                # allowed.
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
                    "Production status: updated"
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

                previous_data_available = (
                    previous_directory
                    is not None
                )

                # ------------------------------------------------
                # MARK RETAINED DATA AS STALE
                #
                # This changes only production_metadata.json.
                # The actual fund data and historical BID history
                # remain untouched.
                # ------------------------------------------------

                if previous_data_available:

                    freshness = (
                        mark_existing_fund_stale(
                            previous_directory,
                            failure_utc,
                        )
                    )

                else:

                    freshness = (
                        build_unavailable_freshness_metadata(
                            failure_utc
                        )
                    )

                failure = build_failure_record(
                    excel_fund,
                    error_text,
                    previous_directory,
                    failure_utc,
                    freshness,
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

                if previous_data_available:

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
                                freshness,

                            "directory":
                                str(
                                    previous_directory
                                ),

                            "reason":
                                error_text,
                        }
                    )

                print()
                print(
                    "============================================================"
                )

                if previous_data_available:

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
                    f"{freshness['status']}"
                )

                print(
                    f"Last successful update: "
                    f"{freshness['lastSuccessfulUpdateUtc']}"
                )

                print(
                    f"Previous data retained: "
                    f"{previous_data_available}"
                )

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
                            run_finished
                            if "run_finished" in locals()
                            else utc_now_iso()
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
    # FINAL RUN TIME
    # ========================================================

    run_finished = utc_now_iso()

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
            "Freshness is based on current-run extraction success "
            "rather than an arbitrary number-of-days threshold.",
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
            unavailable_count,

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

                "carryForwardRawDataAllowed":
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
            unavailable_count,

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
            unavailable_count,

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

                "noCarryForwardRawData":
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

                "emailEnabled":
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
    # FRESHNESS CONSOLE SUMMARY
    # ========================================================

    print()
    print(
        "============================================================"
    )

    print(
        "FRESHNESS SUMMARY"
    )

    print(
        "============================================================"
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
            "Freshness metadata was updated."
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
