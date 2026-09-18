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

    print(f"Saved: {output_file}")


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

    print(f"Saved: {output_file}")


# ============================================================
# EXCEL
# ============================================================

def read_excel_fund():
    print()
    print("========================================")
    print("READING FUNDS LINKS XLSM")
    print("========================================")

    print("Expected file:")
    print(EXCEL_FILE)

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
                "Funds Links.xlsm cell A2 is empty."
            )

        if not pruaccess_name:
            raise ValueError(
                "Funds Links.xlsm cell B2 is empty."
            )

        result = {
            "excelRow": 2,
            "prudentialUrl": str(url).strip(),
            "pruaccessName": str(pruaccess_name).strip(),
        }

        print()
        print("Excel A2 URL:")
        print(result["prudentialUrl"])

        print()
        print("Excel B2 PruAccess name:")
        print(result["pruaccessName"])

        return result

    finally:
        workbook.close()


# ============================================================
# TEXT NORMALISATION
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
# PRUDENTIAL FUND INFORMATION
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
        print("FOUND PRUDENTIAL FUND API:")
        print(response.url)

        print("STATUS:")
        print(response.status)

        try:
            body = await response.text()

            data = json.loads(body)

            captured["url"] = response.url
            captured["status"] = response.status
            captured["data"] = data

        except Exception as exc:
            print(
                "Could not parse Prudential JSON:",
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

    print()
    print("Waiting for Prudential JavaScript...")

    await page.wait_for_timeout(15000)

    if captured["data"] is None:
        raise RuntimeError(
            "Prudential ilpfunds.json response "
            "was not captured."
        )

    data = captured["data"]

    if not isinstance(data, list) or not data:
        raise RuntimeError(
            "Unexpected Prudential fund API response."
        )

    fund = data[0]

    result = {
        "fundName": fund.get("fundName", ""),
        "inceptionDate": fund.get("inceptionDate", ""),
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
    print("Prudential fund name:")
    print(result["fundName"])

    print()
    print("Prudential inception date:")
    print(result["inceptionDate"])

    print()
    print("Current bid price:")
    print(result["bidPrice"])

    print()
    print("Current offer price:")
    print(result["offerPrice"])

    print()
    print("Valuation date:")
    print(result["valuationDate"])

    if not result["inceptionDate"]:
        raise RuntimeError(
            "Prudential did not provide an inceptionDate."
        )

    return result


# ============================================================
# DATE FUNCTIONS
# ============================================================

def convert_prudential_date_to_pruaccess(value):

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
# PRUACCESS INSPECTION
# ============================================================

async def inspect_pruaccess(
    page,
    excel_fund,
    prudential_fund,
):

    print()
    print("========================================")
    print("OPENING PRUACCESS")
    print("========================================")

    captured_requests = []
    captured_responses = []

    async def handle_request(request):

        captured_requests.append({
            "method": request.method,
            "url": request.url,
            "resourceType": request.resource_type,
        })

    async def handle_response(response):

        captured_responses.append({
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

    await page.wait_for_timeout(5000)

    print()
    print("PruAccess loaded:")
    print(page.url)

    # --------------------------------------------------------
    # Save initial page
    # --------------------------------------------------------

    html = await page.content()

    save_text(
        "initial_page.html",
        html,
    )

    visible_text = await page.locator(
        "body"
    ).inner_text()

    save_text(
        "initial_visible_text.txt",
        visible_text,
    )

    # --------------------------------------------------------
    # Inspect fund selector
    # --------------------------------------------------------

    fund_selector = page.locator(
        "#fundName"
    )

    fund_count = await fund_selector.count()

    print()
    print("Fund selector count:")
    print(fund_count)

    if fund_count == 0:
        raise RuntimeError(
            "PruAccess #fundName selector was not found."
        )

    options = await fund_selector.locator(
        "option"
    ).all()

    fund_options = []

    for option in options:

        text = await option.inner_text()
        value = await option.get_attribute("value")

        fund_options.append({
            "text": text.strip(),
            "value": value,
        })

    save_json(
        "fund_options.json",
        fund_options,
    )

    # --------------------------------------------------------
    # Match Excel B2
    # --------------------------------------------------------

    target_name = excel_fund[
        "pruaccessName"
    ]

    target_normalized = normalize_text(
        target_name
    )

    matched_option = None

    for option in fund_options:

        if normalize_text(
            option["text"]
        ) == target_normalized:

            matched_option = option
            break

    print()
    print("Excel B2:")
    print(target_name)

    print()
    print("Matched PruAccess option:")
    print(matched_option)

    if matched_option is None:
        raise RuntimeError(
            "Could not find an exact PruAccess "
            "fund-name match for Excel B2."
        )

    # --------------------------------------------------------
    # Select fund
    # --------------------------------------------------------

    await fund_selector.select_option(
        matched_option["value"]
    )

    print()
    print("Selected PruAccess fund:")
    print(matched_option["text"])

    # --------------------------------------------------------
    # Select TABLE view
    # --------------------------------------------------------

    view_selector = page.locator(
        "#viewType"
    )

    if await view_selector.count() > 0:

        await view_selector.select_option(
            "TBL"
        )

        print()
        print("View type:")
        print("TBL - Table")

    # --------------------------------------------------------
    # DO NOT TOUCH FUND PRICE TYPE
    # --------------------------------------------------------

    price_selector = page.locator(
        "#fundPriceType"
    )

    if await price_selector.count() > 0:

        price_value = await price_selector.input_value()

        print()
        print("Fund price type:")
        print(price_value)

        print(
            "Fund price type was NOT changed."
        )

    # --------------------------------------------------------
    # DATE INPUTS
    # --------------------------------------------------------

    start_input = page.locator(
        'input[name="startDate"]'
    )

    end_input = page.locator(
        'input[name="endDate"]'
    )

    start_count = await start_input.count()
    end_count = await end_input.count()

    print()
    print("Start date input count:")
    print(start_count)

    print()
    print("End date input count:")
    print(end_count)

    current_start = None
    current_end = None

    if start_count > 0:
        current_start = await start_input.first.get_attribute(
            "value"
        )

    if end_count > 0:
        current_end = await end_input.first.get_attribute(
            "value"
        )

    required_start = (
        convert_prudential_date_to_pruaccess(
            prudential_fund["inceptionDate"]
        )
    )

    required_end = get_run_date()

    print()
    print("Current PruAccess start date:")
    print(current_start)

    print()
    print("Current PruAccess end date:")
    print(current_end)

    print()
    print("Required start date:")
    print(required_start)

    print()
    print("Required end date:")
    print(required_end)

    # --------------------------------------------------------
    # Inspect readonly/date-picker structure
    # --------------------------------------------------------

    datepicker_info = await page.evaluate(
        """
        () => {
            const start = document.querySelector(
                'input[name="startDate"]'
            );

            const end = document.querySelector(
                'input[name="endDate"]'
            );

            return {
                start: start ? {
                    id: start.id,
                    name: start.name,
                    value: start.value,
                    readOnly: start.readOnly,
                    disabled: start.disabled,
                    outerHTML: start.outerHTML
                } : null,

                end: end ? {
                    id: end.id,
                    name: end.name,
                    value: end.value,
                    readOnly: end.readOnly,
                    disabled: end.disabled,
                    outerHTML: end.outerHTML
                } : null
            };
        }
        """
    )

    save_json(
        "date_inputs.json",
        datepicker_info,
    )

    # --------------------------------------------------------
    # Inspect forms
    # --------------------------------------------------------

    forms = await page.locator(
        "form"
    ).evaluate_all(
        """
        forms => forms.map(form => ({
            action: form.action,
            method: form.method,
            id: form.id,
            name: form.name,
            outerHTML: form.outerHTML
        }))
        """
    )

    save_json(
        "forms.json",
        forms,
    )

    # --------------------------------------------------------
    # Inspect buttons
    # --------------------------------------------------------

    buttons = await page.locator(
        "button, input[type='submit'], input[type='button'], a"
    ).evaluate_all(
        """
        elements => elements.map(el => ({
            tag: el.tagName,
            type: el.type || null,
            id: el.id || null,
            name: el.name || null,
            text: (el.innerText || el.value || '').trim(),
            href: el.href || null,
            onclick: el.getAttribute('onclick'),
            disabled: el.disabled || false
        }))
        """
    )

    save_json(
        "buttons.json",
        buttons,
    )

    # --------------------------------------------------------
    # Save selected state
    # --------------------------------------------------------

    selected_state = await page.evaluate(
        """
        () => {
            const fund = document.querySelector(
                '#fundName'
            );

            const view = document.querySelector(
                '#viewType'
            );

            const price = document.querySelector(
                '#fundPriceType'
            );

            const start = document.querySelector(
                'input[name="startDate"]'
            );

            const end = document.querySelector(
                'input[name="endDate"]'
            );

            return {
                fund: fund ? {
                    value: fund.value,
                    text: fund.options[
                        fund.selectedIndex
                    ]?.text || ''
                } : null,

                view: view ? {
                    value: view.value,
                    text: view.options[
                        view.selectedIndex
                    ]?.text || ''
                } : null,

                price: price ? {
                    value: price.value,
                    text: price.options[
                        price.selectedIndex
                    ]?.text || ''
                } : null,

                start: start ? start.value : null,
                end: end ? end.value : null
            };
        }
        """
    )

    save_json(
        "selected_state.json",
        selected_state,
    )

    # --------------------------------------------------------
    # Save network information
    # --------------------------------------------------------

    save_json(
        "requests.json",
        captured_requests,
    )

    save_json(
        "responses.json",
        captured_responses,
    )

    # --------------------------------------------------------
    # Save final page before submit
    # --------------------------------------------------------

    save_text(
        "before_submit.html",
        await page.content(),
    )

    save_text(
        "before_submit_visible_text.txt",
        await page.locator(
            "body"
        ).inner_text(),
    )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    summary = {
        "excelFund": excel_fund,
        "prudentialFund": prudential_fund,
        "requiredStartDate": required_start,
        "requiredEndDate": required_end,
        "matchedPruAccessOption": matched_option,
        "selectedState": selected_state,
        "requestCount": len(
            captured_requests
        ),
        "responseCount": len(
            captured_responses
        ),
    }

    save_json(
        "summary.json",
        summary,
    )

    print()
    print("========================================")
    print("PRUACCESS DIAGNOSTIC COMPLETE")
    print("========================================")


# ============================================================
# MAIN
# ============================================================

async def main():

    try:

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

                await inspect_pruaccess(
                    page,
                    excel_fund,
                    prudential_fund,
                )

            finally:

                await browser.close()

    except Exception as exc:

        print()
        print("========================================")
        print("DIAGNOSTIC FAILED")
        print("========================================")

        print(
            type(exc).__name__,
            str(exc),
        )

        save_json(
            "error.json",
            {
                "errorType": type(exc).__name__,
                "error": str(exc),
            },
        )

        raise


if __name__ == "__main__":
    asyncio.run(main())
