import asyncio
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.async_api import async_playwright


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = Path("Funds_Links.xlsm")

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

OUTPUT_DIR = Path("output_pruaccess")
OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# First test only.
# B2 is the first PruAccess fund name.
TEST_EXCEL_ROW = 2


# ============================================================
# READ FUND FROM EXCEL
# ============================================================

def read_excel_fund(row_number):

    print()
    print("========================================")
    print("READING FUNDS_LINKS.XLSM")
    print("========================================")

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_vba=True
    )

    sheet = workbook.active

    url = sheet.cell(
        row=row_number,
        column=1
    ).value

    pruaccess_name = sheet.cell(
        row=row_number,
        column=2
    ).value

    workbook.close()

    if not url:

        raise ValueError(
            f"Column A is empty at row {row_number}"
        )

    if not pruaccess_name:

        raise ValueError(
            f"Column B is empty at row {row_number}"
        )

    result = {
        "excelRow": row_number,
        "sourceUrl": str(url).strip(),
        "pruaccessName": str(
            pruaccess_name
        ).strip()
    }

    print(
        "Excel row:",
        result["excelRow"]
    )

    print(
        "Column A URL:",
        result["sourceUrl"]
    )

    print(
        "Column B PruAccess name:",
        result["pruaccessName"]
    )

    return result


# ============================================================
# MAIN
# ============================================================

async def main():

    # --------------------------------------------------------
    # READ B2
    # --------------------------------------------------------

    excel_fund = read_excel_fund(
        TEST_EXCEL_ROW
    )

    # --------------------------------------------------------
    # END DATE = DATE THE SCRIPT RUNS
    # --------------------------------------------------------

    end_date = datetime.now().strftime(
        "%d-%b-%Y"
    )

    print()
    print("Request end date:")
    print(end_date)

    captured_requests = []
    captured_responses = []

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        context = await browser.new_context(
            accept_downloads=True
        )

        page = await context.new_page()

        # ----------------------------------------------------
        # CAPTURE REQUESTS
        # ----------------------------------------------------

        async def handle_request(request):

            captured_requests.append({
                "method": request.method,
                "url": request.url,
                "resourceType":
                    request.resource_type,
                "postData":
                    request.post_data,
                "headers":
                    dict(request.headers)
            })

            print()
            print("REQUEST:")
            print(
                request.method,
                request.url
            )

            if request.post_data:

                print(
                    "POST DATA:"
                )

                print(
                    request.post_data
                )

        # ----------------------------------------------------
        # CAPTURE RESPONSES
        # ----------------------------------------------------

        async def handle_response(response):

            captured_responses.append({
                "url": response.url,
                "status": response.status,
                "contentType":
                    response.headers.get(
                        "content-type",
                        ""
                    )
            })

            print()
            print("RESPONSE:")
            print(
                response.status,
                response.url
            )

            content_type = (
                response.headers.get(
                    "content-type",
                    ""
                ).lower()
            )

            interesting = any(
                keyword in response.url.lower()
                for keyword in [
                    "fund",
                    "performance",
                    "price",
                    "history",
                    "download",
                    "pdf",
                    "ajax",
                    "json",
                    "api"
                ]
            )

            if interesting:

                try:

                    body = await response.text()

                    filename = (
                        f"response_"
                        f"{len(captured_responses)}.txt"
                    )

                    (
                        OUTPUT_DIR / filename
                    ).write_text(
                        body,
                        encoding="utf-8"
                    )

                    print(
                        "Saved response:",
                        filename
                    )

                except Exception as e:

                    print(
                        "Could not save response:",
                        repr(e)
                    )

        page.on(
            "request",
            handle_request
        )

        page.on(
            "response",
            handle_response
        )

        # ----------------------------------------------------
        # OPEN PRUACCESS
        # ----------------------------------------------------

        print()
        print("========================================")
        print("OPENING PRUACCESS")
        print("========================================")

        await page.goto(
            PRUACCESS_URL,
            wait_until="domcontentloaded",
            timeout=120000
        )

        await page.wait_for_timeout(
            5000
        )

        # ----------------------------------------------------
        # READ ALL PRUACCESS OPTIONS
        # ----------------------------------------------------

        pruaccess_options = await page.evaluate(
            """
            () => {

                const select =
                    document.querySelector(
                        '#fundName'
                    );

                if (!select) {
                    return [];
                }

                return [
                    ...select.options
                ].map(
                    option => ({
                        text:
                            option.text.trim(),
                        value:
                            option.value
                    })
                );
            }
            """
        )

        (
            OUTPUT_DIR /
            "pruaccess_fund_options.json"
        ).write_text(
            json.dumps(
                pruaccess_options,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("========================================")
        print("PRUACCESS FUND OPTIONS")
        print("========================================")

        print(
            "Options found:",
            len(pruaccess_options)
        )

        # ----------------------------------------------------
        # MATCH EXCEL B2 AGAINST PRUACCESS
        # ----------------------------------------------------

        excel_name = (
            excel_fund["pruaccessName"]
        ).strip()

        exact_matches = [
            option
            for option in pruaccess_options
            if option["text"].strip().lower()
            == excel_name.lower()
        ]

        # ----------------------------------------------------
        # If exact match fails, try normalized matching
        # ----------------------------------------------------

        if not exact_matches:

            def normalize(value):

                return " ".join(
                    str(value)
                    .strip()
                    .lower()
                    .split()
                )

            normalized_excel_name = normalize(
                excel_name
            )

            normalized_matches = [
                option
                for option in pruaccess_options
                if normalize(option["text"])
                == normalized_excel_name
            ]

        else:

            normalized_matches = []

        matches = (
            exact_matches
            if exact_matches
            else normalized_matches
        )

        print()
        print("Excel B-column name:")
        print(excel_name)

        print()
        print(
            "Matching PruAccess options:",
            len(matches)
        )

        for match in matches:

            print(
                "MATCH:",
                match
            )

        if not matches:

            print()
            print(
                "ERROR: No matching PruAccess "
                "fund was found."
            )

            # Save diagnostic information.

            (
                OUTPUT_DIR /
                "match_error.json"
            ).write_text(
                json.dumps(
                    {
                        "excelRow":
                            TEST_EXCEL_ROW,
                        "excelName":
                            excel_name,
                        "availableOptions":
                            pruaccess_options
                    },
                    indent=2,
                    ensure_ascii=False
                ),
                encoding="utf-8"
            )

            await browser.close()

            raise ValueError(
                "Could not match Excel "
                "Column B fund name to "
                "PruAccess fund selector."
            )

        if len(matches) > 1:

            print()
            print(
                "WARNING: Multiple matches."
            )

        selected_option = matches[0]

        pruaccess_fund_id = (
            selected_option["value"]
        )

        # ----------------------------------------------------
        # SELECT TABLE
        # ----------------------------------------------------

        await page.locator(
            "#viewType"
        ).select_option(
            "TBL"
        )

        # ----------------------------------------------------
        # SELECT BID PRICE
        # ----------------------------------------------------

        await page.locator(
            "#fundPriceType"
        ).select_option(
            "BID"
        )

        # ----------------------------------------------------
        # SELECT FUND USING EXCEL B2
        # ----------------------------------------------------

        await page.locator(
            "#fundName"
        ).select_option(
            pruaccess_fund_id
        )

        # ----------------------------------------------------
        # GET INCEPTION DATE
        #
        # For now B2's Prudential URL is recorded.
        # The inception date will be discovered from the
        # Prudential fund page in the next stage.
        #
        # For this test we inspect the PruAccess page for
        # any existing/default date and report it.
        # ----------------------------------------------------

        current_start_date = await page.locator(
            "#startDate"
        ).input_value()

        print()
        print(
            "Current PruAccess start date:",
            current_start_date
        )

        # ----------------------------------------------------
        # TEMPORARY START DATE
        #
        # We use the current PruAccess start date for this
        # diagnostic rather than guessing an inception date.
        #
        # Once Prudential inception dates are wired in,
        # this will be replaced automatically.
        # ----------------------------------------------------

        await page.locator(
            "#startDate"
        ).fill(
            current_start_date
        )

        await page.locator(
            "#endDate"
        ).fill(
            end_date
        )

        # ----------------------------------------------------
        # VERIFY SELECTION
        # ----------------------------------------------------

        selected_values = await page.evaluate(
            """
            () => {

                const view =
                    document.querySelector(
                        '#viewType'
                    );

                const price =
                    document.querySelector(
                        '#fundPriceType'
                    );

                const fund =
                    document.querySelector(
                        '#fundName'
                    );

                const start =
                    document.querySelector(
                        '#startDate'
                    );

                const end =
                    document.querySelector(
                        '#endDate'
                    );

                return {

                    viewType:
                        view ? view.value : null,

                    priceType:
                        price ? price.value : null,

                    fundValue:
                        fund ? fund.value : null,

                    fundText:
                        fund
                            ? fund.options[
                                fund.selectedIndex
                            ].text
                            : null,

                    startDate:
                        start
                            ? start.value
                            : null,

                    endDate:
                        end
                            ? end.value
                            : null
                };
            }
            """
        )

        (
            OUTPUT_DIR /
            "selected_values.json"
        ).write_text(
            json.dumps(
                selected_values,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("========================================")
        print("FINAL SELECTION")
        print("========================================")

        print(
            json.dumps(
                selected_values,
                indent=2,
                ensure_ascii=False
            )
        )

        # ----------------------------------------------------
        # SAVE MATCH INFORMATION
        # ----------------------------------------------------

        match_information = {

            "excelRow":
                TEST_EXCEL_ROW,

            "sourceUrl":
                excel_fund["sourceUrl"],

            "excelPruAccessName":
                excel_name,

            "matchedPruAccessName":
                selected_option["text"],

            "pruAccessFundId":
                pruaccess_fund_id,

            "viewType":
                "TBL",

            "priceType":
                "BID",

            "startDate":
                selected_values["startDate"],

            "endDate":
                selected_values["endDate"]
        }

        (
            OUTPUT_DIR /
            "fund_match.json"
        ).write_text(
            json.dumps(
                match_information,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # ----------------------------------------------------
        # FIND SUBMIT BUTTONS
        # ----------------------------------------------------

        print()
        print("========================================")
        print("SUBMIT CONTROLS")
        print("========================================")

        buttons = page.locator(
            'button, '
            'input[type="submit"], '
            'input[type="button"]'
        )

        count = await buttons.count()

        print(
            "Buttons found:",
            count
        )

        for i in range(count):

            button = buttons.nth(i)

            try:

                button_id = (
                    await button.get_attribute(
                        "id"
                    )
                    or ""
                )

                button_name = (
                    await button.get_attribute(
                        "name"
                    )
                    or ""
                )

                button_value = (
                    await button.get_attribute(
                        "value"
                    )
                    or ""
                )

                button_text = (
                    await button.inner_text()
                ).strip()

                print(
                    f"{i}: "
                    f"id={button_id}, "
                    f"name={button_name}, "
                    f"value={button_value}, "
                    f"text={button_text}"
                )

            except Exception:
                pass

        # ----------------------------------------------------
        # SAVE CURRENT PAGE
        # ----------------------------------------------------

        html = await page.content()

        (
            OUTPUT_DIR /
            "before_submit.html"
        ).write_text(
            html,
            encoding="utf-8"
        )

        # ----------------------------------------------------
        # SAVE REQUESTS
        # ----------------------------------------------------

        (
            OUTPUT_DIR /
            "requests.json"
        ).write_text(
            json.dumps(
                captured_requests,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        (
            OUTPUT_DIR /
            "responses.json"
        ).write_text(
            json.dumps(
                captured_responses,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        summary = {

            "excelRow":
                TEST_EXCEL_ROW,

            "excelPruAccessName":
                excel_name,

            "matchedPruAccessName":
                selected_option["text"],

            "pruAccessFundId":
                pruaccess_fund_id,

            "viewType":
                selected_values["viewType"],

            "priceType":
                selected_values["priceType"],

            "startDate":
                selected_values["startDate"],

            "endDate":
                selected_values["endDate"],

            "requestCount":
                len(captured_requests),

            "responseCount":
                len(captured_responses)
        }

        (
            OUTPUT_DIR /
            "summary.json"
        ).write_text(
            json.dumps(
                summary,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("========================================")
        print("PRUACCESS SELECTION TEST COMPLETE")
        print("========================================")

        print(
            "Excel row:",
            TEST_EXCEL_ROW
        )

        print(
            "Excel B name:",
            excel_name
        )

        print(
            "PruAccess match:",
            selected_option["text"]
        )

        print(
            "PruAccess ID:",
            pruaccess_fund_id
        )

        print(
            "View:",
            selected_values["viewType"]
        )

        print(
            "Price:",
            selected_values["priceType"]
        )

        print(
            "Start:",
            selected_values["startDate"]
        )

        print(
            "End:",
            selected_values["endDate"]
        )

        print()
        print(
            "Files saved to:",
            OUTPUT_DIR
        )

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
