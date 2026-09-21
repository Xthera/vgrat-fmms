#!/usr/bin/env python3

"""
VGrat FMS production fund-data builder.

PRODUCTION WORKFLOW
===================

1. Read the Excel master universe from Funds Links.xlsm.
2. Use test_pruaccess.py's already-tested Prudential extraction.
3. Use test_pruaccess.py's already-tested PruAccess extraction.
4. Process every Excel-master fund independently.
5. A fund is published only after that fund passes all validation.
6. A failed fund NEVER overwrites its previous valid production data.
7. Failed funds are flagged in the production run metadata.
8. If one or more funds fail, send ONE Outlook/Microsoft Graph email.
9. If no funds fail, send NO email.
10. Successful funds are still published even if other funds fail.
11. The production dataset therefore remains usable even when
    individual funds fail.

IMPORTANT
=========

The actual extraction and validation rules remain in:

    test_pruaccess.py

This production script intentionally imports those tested functions
rather than rewriting them.

This means:

    test_pruaccess.py
        =
    extraction + validation engine

    build_fund_data.py
        =
    production publication + failure handling + notification

PRODUCTION RULE
===============

PASS:
    New fund data replaces the previous production data for that fund.

FAIL:
    Previous production data is retained.

NO PREVIOUS DATA:
    Fund is flagged as failed and remains absent from the valid
    production fund dataset.

EMAIL
=====

An Outlook/Microsoft Graph failure email is sent ONLY when one or
more funds fail.

No failure:
    No email.

One or more failures:
    One email containing every failed fund.

ENVIRONMENT VARIABLES
=====================

For Outlook/Microsoft Graph email:

    OUTLOOK_TENANT_ID
    OUTLOOK_CLIENT_ID
    OUTLOOK_CLIENT_SECRET
    OUTLOOK_SENDER
    OUTLOOK_RECIPIENTS

OUTLOOK_RECIPIENTS may contain multiple addresses separated by commas.

Example:

    OUTLOOK_RECIPIENTS=you@example.com,backup@example.com

The email is sent through:

    Microsoft Graph
    https://graph.microsoft.com/v1.0/users/{sender}/sendMail

The email system is intentionally independent from the fund-data
publication gate.

If email fails, the fund-data results are NOT rolled back.
The email failure is recorded in the run summary.

EXIT STATUS
===========

0:
    All Excel-master funds passed.

1:
    One or more funds failed.

The production dataset is still updated for every fund that passed.
"""

from __future__ import annotations

import json
import os
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from test_pruaccess import (
    EXCEL_FILE,
    OUTPUT_DIR,
    FUNDS_OUTPUT_DIR,
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
# OUTLOOK / MICROSOFT GRAPH
# ============================================================

GRAPH_TOKEN_URL_TEMPLATE = (
    "https://login.microsoftonline.com/"
    "{tenant_id}/oauth2/v2.0/token"
)

GRAPH_SEND_MAIL_URL_TEMPLATE = (
    "https://graph.microsoft.com/v1.0/users/"
    "{sender}/sendMail"
)

GRAPH_SCOPE = (
    "https://graph.microsoft.com/.default"
)

EMAIL_SUBJECT_PREFIX = (
    "VGrat FMS - Fund Data Update Failure"
)


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

    # Use the same safe_filename logic as the tested
    # extractor, but avoid importing additional internals.
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
# SAVE COMPLETE FUND RESULT TO DIRECTORY
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


# ============================================================
# LOAD PREVIOUS PRODUCTION FUND
# ============================================================

def load_json_file(
    path: Path,
) -> Any:

    return json.loads(
        path.read_text(
            encoding="utf-8"
        )
    )


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

    # --------------------------------------------------------
    # Find all previous production directories for this
    # Excel row. This allows a fund identifier/name to change
    # without leaving stale folders behind.
    # --------------------------------------------------------

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

        # ----------------------------------------------------
        # Move old directories out of the way.
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # Move newly validated fund into production.
        # ----------------------------------------------------

        staging_directory.rename(
            target_directory
        )

    except Exception:

        # ----------------------------------------------------
        # Restore old production data if publication failed.
        # ----------------------------------------------------

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

    # --------------------------------------------------------
    # Publication succeeded.
    # --------------------------------------------------------

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
    The old test extractor writes successful funds to:

        output_pruaccess/funds/

    The new production dataset lives at:

        output_pruaccess/production/funds/

    On the first production run, seed any existing valid legacy
    fund folders into production.

    This allows a newly failed fund to retain the last successful
    test dataset instead of starting with no previous data.
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
# BUILD FAILURE EMAIL
# ============================================================

def build_failure_email(
    run_summary: dict[str, Any],
) -> tuple[str, str]:

    failed_funds = (
        run_summary[
            "failedFunds"
        ]
    )

    subject = (
        f"{EMAIL_SUBJECT_PREFIX}"
        f" ({len(failed_funds)})"
    )

    lines = []

    lines.append(
        "VGrat FMS Fund Data Update"
    )

    lines.append(
        "========================================"
    )

    lines.append(
        f"Run status: {run_summary['status']}"
    )

    lines.append(
        f"Run completed: {run_summary['completedAtUtc']}"
    )

    lines.append(
        ""
    )

    lines.append(
        f"Excel fund universe: "
        f"{run_summary['fundUniverseCount']}"
    )

    lines.append(
        f"Successful updates: "
        f"{run_summary['successfulFundCount']}"
    )

    lines.append(
        f"Failed funds: "
        f"{run_summary['failedFundCount']}"
    )

    lines.append(
        ""
    )

    lines.append(
        "FAILED FUNDS"
    )

    lines.append(
        "----------------------------------------"
    )

    for index, failure in enumerate(
        failed_funds,
        start=1,
    ):

        lines.append(
            f"{index}. "
            f"Excel row: "
            f"{failure['excelRow']}"
        )

        lines.append(
            f"   Fund: "
            f"{failure.get('fundName') or 'Unknown'}"
        )

        lines.append(
            f"   Excel PruAccess name: "
            f"{failure.get('excelPruAccessName') or 'Unknown'}"
        )

        lines.append(
            f"   Fund identifier: "
            f"{failure.get('fundIdentifier') or 'Unknown'}"
        )

        lines.append(
            f"   Fund code: "
            f"{failure.get('fundCode') or 'Unknown'}"
        )

        lines.append(
            f"   Reason: "
            f"{failure['error']}"
        )

        lines.append(
            f"   Previous production data retained: "
            f"{failure['previousProductionDataRetained']}"
        )

        lines.append(
            ""
        )

    lines.append(
        "IMPORTANT"
    )

    lines.append(
        "----------------------------------------"
    )

    lines.append(
        "Successful funds were updated."
    )

    lines.append(
        "Failed funds were not overwritten."
    )

    lines.append(
        "Their previous valid production data was retained "
        "where available."
    )

    return (
        subject,
        "\n".join(
            lines
        ),
    )


# ============================================================
# MICROSOFT GRAPH TOKEN
# ============================================================

def get_graph_access_token() -> str:

    tenant_id = clean_text(
        os.getenv(
            "OUTLOOK_TENANT_ID"
        )
    )

    client_id = clean_text(
        os.getenv(
            "OUTLOOK_CLIENT_ID"
        )
    )

    client_secret = os.getenv(
        "OUTLOOK_CLIENT_SECRET"
    )

    if not tenant_id:

        raise RuntimeError(
            "OUTLOOK_TENANT_ID is not configured."
        )

    if not client_id:

        raise RuntimeError(
            "OUTLOOK_CLIENT_ID is not configured."
        )

    if not client_secret:

        raise RuntimeError(
            "OUTLOOK_CLIENT_SECRET is not configured."
        )

    token_url = (
        GRAPH_TOKEN_URL_TEMPLATE.format(
            tenant_id=tenant_id
        )
    )

    body = urlencode(
        {
            "client_id":
                client_id,

            "client_secret":
                client_secret,

            "scope":
                GRAPH_SCOPE,

            "grant_type":
                "client_credentials",
        }
    ).encode(
        "utf-8"
    )

    request = Request(
        token_url,
        data=body,
        headers={
            "Content-Type":
                "application/x-www-form-urlencoded"
        },
        method="POST",
    )

    try:

        with urlopen(
            request,
            timeout=60,
        ) as response:

            payload = json.loads(
                response.read().decode(
                    "utf-8"
                )
            )

    except HTTPError as error:

        body_text = ""

        try:

            body_text = (
                error.read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )

        except Exception:

            pass

        raise RuntimeError(
            "Microsoft Graph token request failed: "
            f"HTTP {error.code} "
            f"{body_text}"
        ) from error

    except URLError as error:

        raise RuntimeError(
            "Microsoft Graph token request failed: "
            f"{error}"
        ) from error

    token = clean_text(
        payload.get(
            "access_token"
        )
    )

    if not token:

        raise RuntimeError(
            "Microsoft Graph token response did not "
            "contain access_token."
        )

    return token


# ============================================================
# SEND OUTLOOK EMAIL
# ============================================================

def send_outlook_failure_email(
    subject: str,
    body: str,
) -> dict[str, Any]:

    sender = clean_text(
        os.getenv(
            "OUTLOOK_SENDER"
        )
    )

    recipient_text = clean_text(
        os.getenv(
            "OUTLOOK_RECIPIENTS"
        )
    )

    if not sender:

        raise RuntimeError(
            "OUTLOOK_SENDER is not configured."
        )

    if not recipient_text:

        raise RuntimeError(
            "OUTLOOK_RECIPIENTS is not configured."
        )

    recipients = [
        clean_text(
            address
        )
        for address
        in recipient_text.split(",")
        if clean_text(
            address
        )
    ]

    if not recipients:

        raise RuntimeError(
            "OUTLOOK_RECIPIENTS contains no valid recipients."
        )

    token = get_graph_access_token()

    send_url = (
        GRAPH_SEND_MAIL_URL_TEMPLATE.format(
            sender=sender
        )
    )

    payload = {
        "message": {
            "subject":
                subject,

            "body": {
                "contentType":
                    "Text",

                "content":
                    body,
            },

            "toRecipients": [
                {
                    "emailAddress": {
                        "address":
                            address,
                    }
                }
                for address
                in recipients
            ],
        },

        "saveToSentItems":
            True,
    }

    request = Request(
        send_url,
        data=json.dumps(
            payload
        ).encode(
            "utf-8"
        ),
        headers={
            "Authorization":
                f"Bearer {token}",

            "Content-Type":
                "application/json",
        },
        method="POST",
    )

    try:

        with urlopen(
            request,
            timeout=60,
        ) as response:

            status_code = (
                response.status
            )

    except HTTPError as error:

        body_text = ""

        try:

            body_text = (
                error.read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )

        except Exception:

            pass

        raise RuntimeError(
            "Microsoft Graph sendMail failed: "
            f"HTTP {error.code} "
            f"{body_text}"
        ) from error

    except URLError as error:

        raise RuntimeError(
            "Microsoft Graph sendMail failed: "
            f"{error}"
        ) from error

    if status_code != 202:

        raise RuntimeError(
            "Microsoft Graph sendMail returned unexpected "
            f"HTTP status {status_code}."
        )

    return {
        "status":
            "sent",

        "httpStatus":
            status_code,

        "recipientCount":
            len(
                recipients
            ),

        "sender":
            sender,

        "recipients":
            recipients,

        "sentAtUtc":
            utc_now_iso(),
    }


# ============================================================
# LOAD PRODUCTION DATASET
# ============================================================

def collect_production_funds(
    excel_funds: list[dict[str, Any]],
) -> tuple[
    list[dict[str, Any]],
    dict[str, Any],
]:

    production_funds = []

    failed_missing_previous = []

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

            failed_missing_previous.append(
                {
                    "excelRow":
                        row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "excelPruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],
                }
            )

            continue

        try:

            previous = load_previous_fund(
                directory,
                excel_fund,
            )

        except Exception as error:

            failed_missing_previous.append(
                {
                    "excelRow":
                        row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "excelPruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "error":
                        str(error),

                    "directory":
                        str(directory),
                }
            )

            continue

        if previous is None:

            failed_missing_previous.append(
                {
                    "excelRow":
                        row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "excelPruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "directory":
                        str(directory),
                }
            )

            continue

        production_funds.append(
            previous
        )

    return (
        production_funds,
        {
            "missingPreviousData":
                failed_missing_previous
        },
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
    # SEED FIRST PRODUCTION RUN FROM EXISTING TEST OUTPUT
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

    email_status = {
        "required":
            False,

        "status":
            "not_required",
    }

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
                # THIS IS THE PRODUCTION GATE FOR THIS FUND.
                #
                # extract_single_fund() only returns when all
                # windows and full-history validation have passed.
                # ------------------------------------------------

                published_directory = (
                    publish_successful_fund(
                        result,
                        run_staging_directory,
                    )
                )

                result[
                    "productionStatus"
                ] = "updated"

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
                    f"Production directory: "
                    f"{published_directory}"
                )

            except Exception as error:

                error_text = clean_text(
                    str(error)
                )

                # ------------------------------------------------
                # Attempt to retain previous production data.
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

                failure = {
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

                    "failedAtUtc":
                        utc_now_iso(),
                }

                failed_funds.append(
                    failure
                )

                # ------------------------------------------------
                # Save diagnostic failure record.
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
                                fund_name,

                            "fundIdentifier":
                                fund_identifier,

                            "fundCode":
                                fund_code,

                            "status":
                                "retained_previous",

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

                print(
                    "FUND FAILED - PREVIOUS DATA RETAINED"
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
                    f"Previous data retained: "
                    f"{previous_data_available}"
                )

                # ------------------------------------------------
                # IMPORTANT:
                #
                # Do NOT stop the run.
                #
                # Continue processing every remaining fund.
                # ------------------------------------------------

                continue

        await browser.close()

    # ========================================================
    # PRODUCTION DATASET COLLECTION
    # ========================================================

    production_funds = []

    retained_loaded_funds = []

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

            continue

        try:

            production_result = (
                load_previous_fund(
                    directory,
                    excel_fund,
                )
            )

            if production_result:

                production_funds.append(
                    production_result
                )

                if (
                    row
                    not in {
                        item["excelRow"]
                        for item
                        in successful_funds
                    }
                ):

                    retained_loaded_funds.append(
                        production_result
                    )

        except Exception as error:

            print(
                f"WARNING: Could not load production fund "
                f"for Excel row {row}: {error}"
            )

    # ========================================================
    # BUILD CURRENT PRODUCTION BID HISTORY
    # ========================================================

    production_bid_history = {}

    total_observations = 0
    total_windows = 0
    total_pages = 0

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

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "totalHistoricalBidObservations":
            total_observations,

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
            },

        "emailRule":
            {
                "sendOnlyWhenFailures":
                    True,

                "sendWhenNoFailures":
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
                }

                for fund
                in successful_funds
            ],

        "retainedFunds":
            retained_funds,

        "failedFunds":
            failed_funds,
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

        "funds":
            production_funds,

        "failedFunds":
            failed_funds,
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

        "totalHistoricalBidObservations":
            total_observations,

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

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
    # FAILURE EMAIL
    # ========================================================

    if failed_funds:

        email_status[
            "required"
        ] = True

        subject, body = (
            build_failure_email(
                run_summary
            )
        )

        try:

            email_result = (
                send_outlook_failure_email(
                    subject,
                    body,
                )
            )

            email_status.update(
                email_result
            )

            print()
            print(
                "============================================================"
            )

            print(
                "OUTLOOK FAILURE EMAIL SENT"
            )

            print(
                "============================================================"
            )

            print(
                f"Recipients: "
                f"{email_result['recipientCount']}"
            )

        except Exception as error:

            email_status[
                "status"
            ] = "failed"

            email_status[
                "error"
            ] = clean_text(
                str(error)
            )

            print()
            print(
                "============================================================"
            )

            print(
                "WARNING: OUTLOOK FAILURE EMAIL COULD NOT BE SENT"
            )

            print(
                "============================================================"
            )

            print(
                str(error)
            )

            print(
                "Fund production data has NOT been rolled back."
            )

    else:

        print()
        print(
            "No failed funds."
        )

        print(
            "No Outlook email will be sent."
        )

    run_summary[
        "email"
    ] = email_status

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
                f"  Previous data retained: "
                f"{failure['previousProductionDataRetained']}"
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
            "The failed funds were flagged."
        )

        print(
            "The Outlook failure notification was attempted."
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
        "No failure email was sent."
    )

    print(
        "Done."
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    import asyncio

    asyncio.run(
        main()
    )
