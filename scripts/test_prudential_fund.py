#!/usr/bin/env python3

"""
VGrat FMS - Prudential ALL-FUND JSON CAPTURE

MASTER SOURCE
=============

Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    Exact PruAccess fund name


PURPOSE
=======

Extract the official Prudential JSON responses for EVERY fund
listed in Column A.

For each fund, capture:

    - ilpfunds.json
    - ilpseries.json

The complete raw JSON response is saved.

This script does NOT parse individual fund fields.

It is a raw Prudential API capture/extraction layer.

WORKFLOW
========

Funds Links.xlsm
        |
        v
Read every populated URL in Column A
        |
        v
Open one Prudential fund page
        |
        v
Capture official Prudential JSON
        |
        +---- ilpfunds.json
        |
        +---- ilpseries.json
        |
        v
Save raw response
        |
        v
Next fund

OUTPUT
======

output_prudential_json/
    run_summary.json

    funds/
        001_row_2/
            fund_info.json
            response_1.json
            response_2.json
            rendered.html
            visible_text.txt
            requests.json
            summary.json

        002_row_3/
            ...
"""

import asyncio
import json
import re
from pathlib import Path

from openpyxl import load_workbook
from playwright.async_api import async_playwright


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_prudential_json")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

IMPORTANT_PARTS = [
    "ilpfunds.json",
    "ilpseries.json",
]

PAGE_TIMEOUT_MS = 120000

JAVASCRIPT_WAIT_MS = 15000

LAZY_LOAD_WAIT_MS = 5000

DELAY_BETWEEN_FUNDS_MS = 1000


# ============================================================
# HELPERS
# ============================================================

def safe_filename(value):
    """
    Make a string safe for use as a directory/file name.
    """

    value = str(value or "").strip()

    if not value:
        return "unnamed"

    value = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        value
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    value = value.strip(" .")

    if not value:
        return "unnamed"

    return value[:120]


def save_json(path, data):
    """
    Save JSON using UTF-8.
    """

    path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    path.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False
        ),
        encoding="utf-8"
    )


def load_funds_from_excel():
    """
    Read every populated URL from Column A.

    Column A controls the fund universe.

    Column B is retained as the exact PruAccess fund name
    for identification/reference.
    """

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE.resolve()}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    try:

        worksheet = workbook.active

        funds = []

        for row_number in range(
            2,
            worksheet.max_row + 1
        ):

            url_value = worksheet.cell(
                row=row_number,
                column=1
            ).value

            name_value = worksheet.cell(
                row=row_number,
                column=2
            ).value

            if url_value is None:
                continue

            url = str(url_value).strip()

            if not url:
                continue

            fund_name = ""

            if name_value is not None:
                fund_name = str(
                    name_value
                ).strip()

            funds.append({
                "excelRow": row_number,
                "prudentialUrl": url,
                "excelPruAccessName": fund_name
            })

        return funds

    finally:

        workbook.close()


# ============================================================
# PROCESS ONE FUND
# ============================================================

async def process_fund(
    browser,
    fund,
    fund_number,
    total_funds
):
    """
    Extract all matching Prudential JSON responses for
    one fund.
    """

    excel_row = fund["excelRow"]

    prudential_url = fund[
        "prudentialUrl"
    ]

    excel_pruaccess_name = fund[
        "excelPruAccessName"
    ]

    # --------------------------------------------------------
    # Fund output directory
    # --------------------------------------------------------

    name_part = safe_filename(
        excel_pruaccess_name
    )

    fund_dir = (
        FUNDS_OUTPUT_DIR
        / f"{fund_number:03d}_row_{excel_row}_{name_part}"
    )

    fund_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print("=" * 80)
    print(
        f"FUND {fund_number}/{total_funds}"
    )
    print(
        f"Excel row: {excel_row}"
    )
    print(
        f"Fund name: "
        f"{excel_pruaccess_name or '(Column B blank)'}"
    )
    print(
        f"URL: {prudential_url}"
    )
    print("=" * 80)

    # --------------------------------------------------------
    # Save fund information
    # --------------------------------------------------------

    save_json(
        fund_dir / "fund_info.json",
        {
            "fundNumber": fund_number,
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pruaccess_name
        }
    )

    captured_json = []

    captured_requests = []

    # --------------------------------------------------------
    # Create browser context
    # --------------------------------------------------------

    context = await browser.new_context(
        viewport={
            "width": 1440,
            "height": 1000
        }
    )

    page = await context.new_page()

    # --------------------------------------------------------
    # Response handler
    # --------------------------------------------------------

    async def handle_response(response):

        url = response.url

        if not any(
            part in url
            for part in IMPORTANT_PARTS
        ):
            return

        print()
        print(
            "FOUND PRUDENTIAL JSON:"
        )
        print(url)
        print(
            "STATUS:",
            response.status
        )

        try:

            body = await response.text()

            response_number = (
                len(captured_json) + 1
            )

            filename = (
                f"response_{response_number}.json"
            )

            output_file = (
                fund_dir / filename
            )

            output_file.write_text(
                body,
                encoding="utf-8"
            )

            captured_json.append({
                "filename": filename,
                "url": url,
                "status": response.status,
                "contentType": (
                    response.headers.get(
                        "content-type",
                        ""
                    )
                )
            })

        except Exception as e:

            print(
                "Could not read response:",
                repr(e)
            )

    # --------------------------------------------------------
    # Request handler
    # --------------------------------------------------------

    async def handle_request(request):

        url = request.url

        if any(
            part in url
            for part in IMPORTANT_PARTS
        ):

            captured_requests.append({
                "method": request.method,
                "url": request.url,
                "postData": request.post_data
            })

    page.on(
        "response",
        handle_response
    )

    page.on(
        "request",
        handle_request
    )

    # --------------------------------------------------------
    # Open Prudential page
    # --------------------------------------------------------

    try:

        print()
        print(
            "Opening Prudential page..."
        )

        try:

            await page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS
            )

        except Exception as e:

            print()
            print(
                "PAGE NAVIGATION ERROR:"
            )
            print(
                repr(e)
            )

            # Continue because the page may have already
            # triggered useful network requests.

        print()
        print(
            "Waiting for Prudential JavaScript..."
        )

        await page.wait_for_timeout(
            JAVASCRIPT_WAIT_MS
        )

        # ----------------------------------------------------
        # Scroll page
        # ----------------------------------------------------

        try:

            await page.evaluate(
                """
                window.scrollTo(
                    0,
                    document.body.scrollHeight
                );
                """
            )

        except Exception as e:

            print(
                "Scroll warning:",
                repr(e)
            )

        await page.wait_for_timeout(
            LAZY_LOAD_WAIT_MS
        )

        # ----------------------------------------------------
        # Save rendered HTML
        # ----------------------------------------------------

        try:

            html = await page.content()

            (
                fund_dir / "rendered.html"
            ).write_text(
                html,
                encoding="utf-8"
            )

        except Exception as e:

            print(
                "Could not save rendered HTML:",
                repr(e)
            )

        # ----------------------------------------------------
        # Save visible text
        # ----------------------------------------------------

        try:

            visible_text = await page.locator(
                "body"
            ).inner_text()

            (
                fund_dir / "visible_text.txt"
            ).write_text(
                visible_text,
                encoding="utf-8"
            )

        except Exception as e:

            print(
                "Could not save visible text:",
                repr(e)
            )

        # ----------------------------------------------------
        # Save request list
        # ----------------------------------------------------

        save_json(
            fund_dir / "requests.json",
            captured_requests
        )

        # ----------------------------------------------------
        # Save summary
        # ----------------------------------------------------

        summary = {
            "fundNumber": fund_number,
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": (
                excel_pruaccess_name
            ),
            "jsonResponsesCaptured": (
                len(captured_json)
            ),
            "responses": captured_json
        }

        save_json(
            fund_dir / "summary.json",
            summary
        )

        # ----------------------------------------------------
        # Determine result
        # ----------------------------------------------------

        has_prudential_json = (
            len(captured_json) > 0
        )

        print()
        print(
            "----------------------------------------"
        )

        print(
            "JSON responses captured:",
            len(captured_json)
        )

        if has_prudential_json:

            print(
                "STATUS: SUCCESS"
            )

        else:

            print(
                "STATUS: NO PRUDENTIAL JSON CAPTURED"
            )

        print(
            "----------------------------------------"
        )

        return {
            "status": (
                "success"
                if has_prudential_json
                else "no_json"
            ),
            "fundNumber": fund_number,
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": (
                excel_pruaccess_name
            ),
            "jsonResponsesCaptured": (
                len(captured_json)
            ),
            "outputDirectory": str(
                fund_dir
            ),
            "responses": captured_json
        }

    except Exception as e:

        print()
        print(
            "FUND EXTRACTION ERROR:"
        )
        print(
            repr(e)
        )

        return {
            "status": "failed",
            "fundNumber": fund_number,
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": (
                excel_pruaccess_name
            ),
            "error": repr(e),
            "outputDirectory": str(
                fund_dir
            )
        }

    finally:

        await page.close()
        await context.close()


# ============================================================
# MAIN
# ============================================================

async def main():

    print()
    print("=" * 80)
    print(
        "VGrat FMS - PRUDENTIAL ALL-FUND JSON CAPTURE"
    )
    print("=" * 80)

    # --------------------------------------------------------
    # Prepare output directories
    # --------------------------------------------------------

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Read Excel
    # --------------------------------------------------------

    print()
    print(
        "Reading:",
        EXCEL_FILE
    )

    funds = load_funds_from_excel()

    total_funds = len(funds)

    print()
    print(
        "Funds found in Column A:",
        total_funds
    )

    if total_funds == 0:

        print()
        print(
            "ERROR: No populated fund URLs found."
        )

        return

    # --------------------------------------------------------
    # Results
    # --------------------------------------------------------

    results = []

    successful = 0
    no_json = 0
    failed = 0

    # --------------------------------------------------------
    # Launch browser
    # --------------------------------------------------------

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        try:

            # ------------------------------------------------
            # Process funds ONE AT A TIME
            # ------------------------------------------------

            for fund_number, fund in enumerate(
                funds,
                start=1
            ):

                result = await process_fund(
                    browser=browser,
                    fund=fund,
                    fund_number=fund_number,
                    total_funds=total_funds
                )

                results.append(
                    result
                )

                if result["status"] == "success":

                    successful += 1

                elif result["status"] == "no_json":

                    no_json += 1

                else:

                    failed += 1

                # --------------------------------------------
                # Save progress after every fund
                # --------------------------------------------

                progress_summary = {
                    "status": "running",
                    "totalFunds": total_funds,
                    "processedFunds": len(
                        results
                    ),
                    "successfulFunds": successful,
                    "noJsonFunds": no_json,
                    "failedFunds": failed,
                    "results": results
                }

                save_json(
                    OUTPUT_DIR
                    / "run_summary.json",
                    progress_summary
                )

                # --------------------------------------------
                # Delay before next fund
                # --------------------------------------------

                if fund_number < total_funds:

                    await asyncio.sleep(
                        DELAY_BETWEEN_FUNDS_MS
                        / 1000
                    )

        finally:

            await browser.close()

    # --------------------------------------------------------
    # Final summary
    # --------------------------------------------------------

    final_summary = {
        "status": (
            "completed"
            if failed == 0
            else "completed_with_failures"
        ),
        "totalFunds": total_funds,
        "processedFunds": len(
            results
        ),
        "successfulFunds": successful,
        "noJsonFunds": no_json,
        "failedFunds": failed,
        "results": results
    }

    save_json(
        OUTPUT_DIR
        / "run_summary.json",
        final_summary
    )

    # --------------------------------------------------------
    # Final console output
    # --------------------------------------------------------

    print()
    print("=" * 80)
    print(
        "PRUDENTIAL ALL-FUND JSON CAPTURE COMPLETE"
    )
    print("=" * 80)

    print(
        "Total funds:",
        total_funds
    )

    print(
        "Successful:",
        successful
    )

    print(
        "No JSON:",
        no_json
    )

    print(
        "Failed:",
        failed
    )

    print()
    print(
        "Output directory:"
    )

    print(
        OUTPUT_DIR.resolve()
    )

    print()
    print(
        "Run summary:"
    )

    print(
        (
            OUTPUT_DIR
            / "run_summary.json"
        ).resolve()
    )

    print("=" * 80)


if __name__ == "__main__":
    asyncio.run(main())
