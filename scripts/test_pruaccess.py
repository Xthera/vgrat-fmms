#!/usr/bin/env python3

import asyncio
import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.async_api import async_playwright


# ============================================================
# PATHS
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent

EXCEL_FILE = REPO_DIR / "Funds Links.xlsm"

OUTPUT_DIR = REPO_DIR / "output_pruaccess"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)


# ============================================================
# SAVE HELPERS
# ============================================================

def save_text(filename, text):

    output_file = OUTPUT_DIR / filename

    output_file.write_text(
        str(text),
        encoding="utf-8",
    )

    print("Saved:", output_file)


def save_json(filename, data):

    output_file = OUTPUT_DIR / filename

    output_file.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
            default=str,
        ),
        encoding="utf-8",
    )

    print("Saved:", output_file)


# ============================================================
# EXCEL
# ============================================================

def read_excel_fund():

    print()
    print("========================================")
    print("READING FUNDS LINKS XLSM")
    print("========================================")

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        filename=EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_vba=True,
    )

    try:

        sheet = workbook.active

        url = sheet["A2"].value
        pruaccess_name = sheet["B2"].value

        if not url:
            raise ValueError(
                "Funds Links.xlsm A2 is empty."
            )

        if not pruaccess_name:
            raise ValueError(
                "Funds Links.xlsm B2 is empty."
            )

        result = {
            "excelRow": 2,
            "prudentialUrl": str(url).strip(),
            "pruaccessName": str(
                pruaccess_name
            ).strip(),
        }

        print()
        print("Excel A2:")
        print(result["prudentialUrl"])

        print()
        print("Excel B2:")
        print(result["pruaccessName"])

        return result

    finally:

        workbook.close()


# ============================================================
# NORMALISE TEXT
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    text = str(value).strip()

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.casefold()


# ============================================================
# PRUDENTIAL FUND API
# ============================================================

async def get_prudential_fund(
    page,
    prudential_url,
):

    print()
    print("========================================")
    print("GETTING PRUDENTIAL FUND INFORMATION")
    print("========================================")

    captured = {
        "url": None,
        "status": None,
        "data": None,
    }

    async def handle_response(response):

        if "ilpfunds.json" not in response.url:
            return

        print()
        print("FOUND PRUDENTIAL API:")
        print(response.url)

        try:

            body = await response.text()

            data = json.loads(body)

            captured["url"] = response.url
            captured["status"] = response.status
            captured["data"] = data

        except Exception as exc:

            print(
                "JSON parse error:",
                repr(exc),
            )

    page.on(
        "response",
        handle_response,
    )

    await page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    await page.wait_for_timeout(
        15000
    )

    if captured["data"] is None:

        raise RuntimeError(
            "Prudential ilpfunds.json "
            "was not captured."
        )

    if not isinstance(
        captured["data"],
        list,
    ):

        raise RuntimeError(
            "Unexpected Prudential API structure."
        )

    if not captured["data"]:

        raise RuntimeError(
            "Prudential API returned no fund."
        )

    fund = captured["data"][0]

    result = {
        "fundName": fund.get(
            "fundName",
            "",
        ),
        "inceptionDate": fund.get(
            "inceptionDate",
            "",
        ),
        "fundIdentifier": fund.get(
            "fundIdentifier",
            "",
        ),
        "fundCode": fund.get(
            "fundCode",
            "",
        ),
        "bidPrice": fund.get(
            "bidPrice",
            "",
        ),
        "offerPrice": fund.get(
            "offerPrice",
            "",
        ),
        "valuationDate": fund.get(
            "valuationDate",
            "",
        ),
        "raw": fund,
    }

    print()
    print("Prudential fund:")
    print(result["fundName"])

    print()
    print("Inception:")
    print(result["inceptionDate"])

    print()
    print("Bid:")
    print(result["bidPrice"])

    print()
    print("Valuation:")
    print(result["valuationDate"])

    return result


# ============================================================
# DATE CONVERSION
# ============================================================

def convert_prudential_date(
    value,
):

    parsed = datetime.strptime(
        value,
        "%d/%m/%Y",
    )

    return parsed.strftime(
        "%d-%b-%Y"
    )


def get_run_date():

    return datetime.now().strftime(
        "%d-%b-%Y"
    )


# ============================================================
# PRUACCESS
# ============================================================

async def run_pruaccess(
    page,
    excel_fund,
    prudential_fund,
):

    print()
    print("========================================")
    print("OPENING PRUACCESS")
    print("========================================")

    requests = []
    responses = []

    async def handle_request(request):

        requests.append({
            "method": request.method,
            "url": request.url,
            "resourceType": request.resource_type,
        })

    async def handle_response(response):

        responses.append({
            "status": response.status,
            "url": response.url,
        })

    page.on(
        "request",
        handle_request,
    )

    page.on(
        "response",
        handle_response,
    )

    await page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    await page.wait_for_timeout(
        5000
    )

    # --------------------------------------------------------
    # FUND OPTIONS
    # --------------------------------------------------------

    fund_selector = page.locator(
        "#fundName"
    )

    options = await fund_selector.locator(
        "option"
    ).all()

    fund_options = []

    for option in options:

        text = await option.inner_text()

        value = await option.get_attribute(
            "value"
        )

        fund_options.append({
            "text": text.strip(),
            "value": value,
        })

    save_json(
        "fund_options.json",
        fund_options,
    )

    # --------------------------------------------------------
    # MATCH B2
    # --------------------------------------------------------

    target = normalize_text(
        excel_fund["pruaccessName"]
    )

    matched = None

    for option in fund_options:

        if normalize_text(
            option["text"]
        ) == target:

            matched = option
            break

    print()
    print("Excel B2:")
    print(
        excel_fund["pruaccessName"]
    )

    print()
    print("Matched option:")
    print(matched)

    if matched is None:

        raise RuntimeError(
            "No exact PruAccess fund match."
        )

    # --------------------------------------------------------
    # SELECT FUND
    # --------------------------------------------------------

    await fund_selector.select_option(
        matched["value"]
    )

    # --------------------------------------------------------
    # TABLE VIEW
    # --------------------------------------------------------

    await page.locator(
        "#viewType"
    ).select_option(
        "TBL"
    )

    # --------------------------------------------------------
    # FUND PRICE TYPE
    # --------------------------------------------------------
    #
    # IMPORTANT:
    # This field is disabled/greyed out.
    # We deliberately DO NOT modify it.
    #

    price_selector = page.locator(
        "#fundPriceType"
    )

    price_value = None

    if await price_selector.count():

        price_value = await price_selector.input_value()

    print()
    print("Fund price type:")
    print(price_value)

    # --------------------------------------------------------
    # REQUIRED DATES
    # --------------------------------------------------------

    required_start = (
        convert_prudential_date(
            prudential_fund[
                "inceptionDate"
            ]
        )
    )

    required_end = get_run_date()

    print()
    print("Required start:")
    print(required_start)

    print()
    print("Required end:")
    print(required_end)

    # --------------------------------------------------------
    # FORM FIELDS
    # --------------------------------------------------------

    csrf = await page.locator(
        'input[name="_csrf"]'
    ).input_value()

    selected_funds = matched["value"]

    # --------------------------------------------------------
    # SAVE PRE-SUBMIT STATE
    # --------------------------------------------------------

    pre_submit = {
        "csrfPresent": bool(csrf),
        "selectedFunds": selected_funds,
        "selectedFundName": matched["text"],
        "viewType": "TBL",
        "fundPriceType": price_value,
        "startDate": required_start,
        "endDate": required_end,
    }

    save_json(
        "pre_submit.json",
        pre_submit,
    )

    # --------------------------------------------------------
    # SUBMIT FORM
    # --------------------------------------------------------

    print()
    print("========================================")
    print("SUBMITTING PRUACCESS FORM")
    print("========================================")

    async with page.expect_navigation(
        wait_until="domcontentloaded",
        timeout=120000,
    ):

        await page.locator(
            "#fundForm"
        ).evaluate(
            """
            (form, data) => {

                const setField = (
                    name,
                    value
                ) => {

                    let field =
                        form.querySelector(
                            `[name="${name}"]`
                        );

                    if (!field) {

                        field =
                            document.createElement(
                                "input"
                            );

                        field.type = "hidden";
                        field.name = name;

                        form.appendChild(
                            field
                        );
                    }

                    field.value = value;
                };

                setField(
                    "_csrf",
                    data.csrf
                );

                setField(
                    "viewType",
                    "TBL"
                );

                setField(
                    "selectedFunds",
                    data.selectedFunds
                );

                setField(
                    "startDate",
                    data.startDate
                );

                setField(
                    "endDate",
                    data.endDate
                );

                form.submit();
            }
            """,
            {
                "csrf": csrf,
                "selectedFunds": selected_funds,
                "startDate": required_start,
                "endDate": required_end,
            },
        )

    # --------------------------------------------------------
    # WAIT FOR RESULT
    # --------------------------------------------------------

    await page.wait_for_timeout(
        5000
    )

    print()
    print("Returned URL:")
    print(page.url)

    # --------------------------------------------------------
    # SAVE RESULT
    # --------------------------------------------------------

    result_html = await page.content()

    result_text = await page.locator(
        "body"
    ).inner_text()

    save_text(
        "result.html",
        result_html,
    )

    save_text(
        "result_visible_text.txt",
        result_text,
    )

    save_json(
        "requests.json",
        requests,
    )

    save_json(
        "responses.json",
        responses,
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    summary = {
        "excelFund": excel_fund,
        "prudentialFund": prudential_fund,
        "matchedPruAccessFund": matched,
        "startDate": required_start,
        "endDate": required_end,
        "viewType": "TBL",
        "fundPriceType": price_value,
        "returnedUrl": page.url,
        "requestCount": len(requests),
        "responseCount": len(responses),
    }

    save_json(
        "summary.json",
        summary,
    )

    print()
    print("========================================")
    print("PRUACCESS SEARCH COMPLETE")
    print("========================================")


# ============================================================
# MAIN
# ============================================================

async def main():

    excel_fund = read_excel_fund()

    save_json(
        "excel_fund.json",
        excel_fund,
    )

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True,
        )

        try:

            page = await browser.new_page(
                viewport={
                    "width": 1440,
                    "height": 1000,
                }
            )

            prudential_fund = (
                await get_prudential_fund(
                    page,
                    excel_fund[
                        "prudentialUrl"
                    ],
                )
            )

            save_json(
                "prudential_fund.json",
                prudential_fund,
            )

            await run_pruaccess(
                page,
                excel_fund,
                prudential_fund,
            )

        finally:

            await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
