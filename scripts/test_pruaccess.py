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
# EXCEL
# ============================================================

def read_excel_fund():
    """
    Read the first fund from Funds_Links.xlsm.

    Column A:
        Prudential fund URL

    Column B:
        PruAccess fund name

    B2 is used for the PruAccess fund selection.
    """

    print()
    print("========================================")
    print("READING FUNDS_LINKS.XLSM")
    print("========================================")
    print("Expected file:")
    print(EXCEL_FILE)
    print()

    if not EXCEL_FILE.exists():
        print("Repository contents:")
        for item in REPO_DIR.iterdir():
            print(" -", item.name)

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

        print("Worksheet:", sheet.title)

        url = sheet["A2"].value
        pruaccess_name = sheet["B2"].value

        if not url:
            raise ValueError(
                "Funds_Links.xlsm cell A2 is empty."
            )

        if not pruaccess_name:
            raise ValueError(
                "Funds_Links.xlsm cell B2 is empty."
            )

        url = str(url).strip()
        pruaccess_name = str(pruaccess_name).strip()

        print()
        print("Excel A2 URL:")
        print(url)

        print()
        print("Excel B2 PruAccess name:")
        print(pruaccess_name)

        return {
            "excelRow": 2,
            "prudentialUrl": url,
            "pruaccessName": pruaccess_name,
        }

    finally:
        workbook.close()


# ============================================================
# TEXT NORMALISATION
# ============================================================

def normalize_text(value):
    """
    Normalise text so Excel B2 can be matched against
    PruAccess option text safely.
    """

    if value is None:
        return ""

    text = str(value).strip()

    text = re.sub(r"\s+", " ", text)

    return text.casefold()


# ============================================================
# PRUDENTIAL INCEPTION DATE
# ============================================================

async def get_prudential_inception_date(page, prudential_url):
    """
    Open the Prudential fund page and capture the ilpfunds.json
    response.

    The inceptionDate from Prudential is used as the PruAccess
    start date.
    """

    print()
    print("========================================")
    print("GETTING PRUDENTIAL FUND INFORMATION")
    print("========================================")

    print("Prudential URL:")
    print(prudential_url)

    captured = {
        "url": None,
        "status": None,
        "data": None,
    }

    async def handle_response(response):
        response_url = response.url

        if "ilpfunds.json" not in response_url:
            return

        print()
        print("FOUND PRUDENTIAL FUND API:")
        print(response_url)
        print("STATUS:", response.status)

        try:
            body = await response.text()

            data = json.loads(body)

            captured["url"] = response_url
            captured["status"] = response.status
            captured["data"] = data

        except Exception as exc:
            print(
                "Could not parse Prudential JSON:",
                repr(exc)
            )

    page.on("response", handle_response)

    await page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    await page.wait_for_timeout(15000)

    if captured["data"] is None:
        raise RuntimeError(
            "Prudential ilpfunds.json response was not captured."
        )

    data = captured["data"]

    if not isinstance(data, list) or not data:
        raise RuntimeError(
            "Unexpected Prudential ilpfunds.json structure."
        )

    fund = data[0]

    fund_name = fund.get("fundName", "")
    inception_date = fund.get("inceptionDate", "")

    print()
    print("Prudential fund name:")
    print(fund_name)

    print()
    print("Prudential inception date:")
    print(inception_date)

    if not inception_date:
        raise RuntimeError(
            "Prudential did not provide an inceptionDate."
        )

    return {
        "fundName": fund_name,
        "inceptionDate": inception_date,
        "fundIdentifier": fund.get("fundIdentifier", ""),
        "fundCode": fund.get("fundCode", ""),
        "bidPrice": fund.get("bidPrice", ""),
        "offerPrice": fund.get("offerPrice", ""),
        "valuationDate": fund.get("valuationDate", ""),
        "raw": fund,
    }


# ============================================================
# DATE CONVERSION
# ============================================================

def convert_prudential_date_to_pruaccess(value):
    """
    Convert:

        03/11/2021

    into:

        03-Nov-2021
    """

    parsed = datetime.strptime(
        value,
        "%d/%m/%Y",
    )

    return parsed.strftime("%d-%b-%Y")


def get_run_date():
    """
    Use the date on which the GitHub Actions request is made.
    """

    return datetime.now().strftime("%d-%b-%Y")


# ============================================================
# SAVE JSON
# ============================================================

def save_json(filename, data):
    output_file = OUTPUT_DIR / filename

    output_file.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print("Saved:", output_file)


# ============================================================
# MAIN
# ============================================================

async def main():

    # --------------------------------------------------------
    # 1. Read Excel
    # --------------------------------------------------------

    excel_fund = read_excel_fund()

    save_json(
        "excel_fund.json",
        excel_fund,
    )

    # --------------------------------------------------------
    # 2. Start browser
    # --------------------------------------------------------

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        page = await browser.new_page(
            viewport={
                "width": 1440,
                "height": 1000,
            }
        )

        # ----------------------------------------------------
        # 3. Get Prudential inception date
        # ----------------------------------------------------

        prudential_fund = (
            await get_prudential_inception_date(
                page,
                excel_fund["prudentialUrl"],
            )
        )

        save_json(
            "prudential_fund.json",
            prudential_fund,
        )

        # ----------------------------------------------------
        # 4. Convert dates
        # ----------------------------------------------------

        start_date = (
            convert_prudential_date_to_pruaccess(
                prudential_fund["inceptionDate"]
            )
        )

        end_date = get_run_date()

        print()
        print("========================================")
        print("DATES")
        print("========================================")
        print("Inception date:", prudential_fund["inceptionDate"])
        print("PruAccess start:", start_date)
        print("PruAccess end:", end_date)

        # ----------------------------------------------------
        # 5. Open PruAccess
        # ----------------------------------------------------

        print()
        print("========================================")
        print("OPENING PRUACCESS")
        print("========================================")
        print(PRUACCESS_URL)

        captured_requests = []
        captured_responses = []

        async def handle_request(request):
            captured_requests.append({
                "method": request.method,
                "url": request.url,
                "postData": request.post_data,
            })

        async def handle_response(response):
            try:
                captured_responses.append({
                    "status": response.status,
                    "url": response.url,
                    "contentType": response.headers.get(
                        "content-type",
                        "",
                    ),
                })
            except Exception:
                pass

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

        # ----------------------------------------------------
        # 6. Read PruAccess fund selector
        # ----------------------------------------------------

        print()
        print("========================================")
        print("READING PRUACCESS FUND OPTIONS")
        print("========================================")

        fund_select = page.locator("#fundName")

        await fund_select.wait_for(
            state="visible",
            timeout=30000,
        )

        options = await fund_select.locator(
            "option"
        ).evaluate_all(
            """
            options => options.map(option => ({
                text: option.textContent.trim(),
                value: option.value
            }))
            """
        )

        print(
            "PruAccess fund options:",
            len(options)
        )

        save_json(
            "pruaccess_fund_options.json",
            options,
        )

        # ----------------------------------------------------
        # 7. Match Excel B2 to PruAccess
        # ----------------------------------------------------

        excel_name = excel_fund["pruaccessName"]
        normalized_excel_name = normalize_text(
            excel_name
        )

        exact_matches = []

        for option in options:

            option_text = str(
                option.get("text", "")
            ).strip()

            if normalize_text(option_text) == normalized_excel_name:
                exact_matches.append(option)

        print()
        print("Excel B2:")
        print(excel_name)

        print()
        print("Exact matches:")
        print(len(exact_matches))

        if not exact_matches:

            print()
            print("NO EXACT MATCH FOUND.")

            # Also try a contains comparison purely for
            # diagnostics. We do NOT automatically select it.
            possible_matches = []

            for option in options:

                option_text = str(
                    option.get("text", "")
                ).strip()

                if (
                    normalized_excel_name
                    in normalize_text(option_text)
                    or normalize_text(option_text)
                    in normalized_excel_name
                ):
                    possible_matches.append(option)

            save_json(
                "possible_fund_matches.json",
                possible_matches,
            )

            raise RuntimeError(
                "Excel B2 fund name does not exactly match "
                "any PruAccess fund option."
            )

        if len(exact_matches) > 1:
            raise RuntimeError(
                "More than one PruAccess option exactly "
                "matches Excel B2."
            )

        matched_option = exact_matches[0]

        print()
        print("MATCH FOUND")
        print("Text:", matched_option["text"])
        print("Value:", matched_option["value"])

        # ----------------------------------------------------
        # 8. Select fund
        # ----------------------------------------------------

        await fund_select.select_option(
            matched_option["value"]
        )

        # ----------------------------------------------------
        # 9. Select Table
        # ----------------------------------------------------

        print()
        print("Selecting View Type = Table")

        await page.locator(
            "#viewType"
        ).select_option("TBL")

       # ----------------------------------------------------
# Set dates using the actual date-picker inputs
# ----------------------------------------------------

print("Setting Start Date =", start_date)

start_input = page.locator(
    'input[name="startDate"]'
)

end_input = page.locator(
    'input[name="endDate"]'
)

print("Current PruAccess start date:")
print(await start_input.input_value())

print("Current PruAccess end date:")
print(await end_input.input_value())

# The inputs are readonly, so we do NOT use fill().
# We will inspect the date-picker implementation first.
      
        # ----------------------------------------------------
        # 13. Save selected values
        # ----------------------------------------------------

        selected_values = await page.evaluate(
            """
            () => ({
                viewType:
                    document.querySelector("#viewType")?.value,

                fundPriceType:
                    document.querySelector("#fundPriceType")?.value,

                fundName:
                    document.querySelector("#fundName")?.value,

                startDate:
                    document.querySelector("#startDate")?.value,

                endDate:
                    document.querySelector("#endDate")?.value,

                csrf:
                    document.querySelector(
                        'input[name="_csrf"]'
                    )?.value,

                selectedFunds:
                    document.querySelector(
                        'input[name="_selectedFunds"]'
                    )?.value
            })
            """
        )

        print()
        print("========================================")
        print("FINAL SELECTED VALUES")
        print("========================================")

        print(
            json.dumps(
                selected_values,
                indent=2,
                ensure_ascii=False,
            )
        )

        save_json(
            "selected_values.json",
            selected_values,
        )

        # ----------------------------------------------------
        # 14. Save current HTML
        # ----------------------------------------------------

        html = await page.content()

        (
            OUTPUT_DIR / "before_submit.html"
        ).write_text(
            html,
            encoding="utf-8",
        )

        visible_text = await page.locator(
            "body"
        ).inner_text()

        (
            OUTPUT_DIR / "before_submit_visible_text.txt"
        ).write_text(
            visible_text,
            encoding="utf-8",
        )

        # ----------------------------------------------------
        # 15. Inspect forms
        # ----------------------------------------------------

        forms = await page.evaluate(
            """
            () => Array.from(
                document.forms
            ).map((form, index) => ({
                index,
                action: form.action,
                method: form.method,
                id: form.id,
                name: form.name,
                target: form.target,
                inputs: Array.from(
                    form.querySelectorAll("input")
                ).map(input => ({
                    name: input.name,
                    type: input.type,
                    value: input.value,
                    id: input.id
                })),
                selects: Array.from(
                    form.querySelectorAll("select")
                ).map(select => ({
                    name: select.name,
                    id: select.id,
                    value: select.value
                })),
                buttons: Array.from(
                    form.querySelectorAll(
                        "button, input[type=submit]"
                    )
                ).map(button => ({
                    tag: button.tagName,
                    type: button.type,
                    name: button.name,
                    value: button.value,
                    id: button.id,
                    text: button.innerText || ""
                }))
            }))
            """
        )

        save_json(
            "forms.json",
            forms,
        )

        # ----------------------------------------------------
        # 16. Save network diagnostics
        # ----------------------------------------------------

        save_json(
            "requests.json",
            captured_requests,
        )

        save_json(
            "responses.json",
            captured_responses,
        )

        # ----------------------------------------------------
        # 17. Save summary
        # ----------------------------------------------------

        summary = {
            "excelFile": str(EXCEL_FILE),
            "excelRow": excel_fund["excelRow"],
            "excelPruAccessName": excel_name,
            "prudentialFundName": prudential_fund["fundName"],
            "fundIdentifier": prudential_fund[
                "fundIdentifier"
            ],
            "fundCode": prudential_fund[
                "fundCode"
            ],
            "inceptionDate": prudential_fund[
                "inceptionDate"
            ],
            "pruaccessStartDate": start_date,
            "pruaccessEndDate": end_date,
            "selectedViewType": selected_values[
                "viewType"
            ],
            "selectedPriceType": selected_values[
                "fundPriceType"
            ],
            "selectedFundValue": selected_values[
                "fundName"
            ],
            "matchedPruAccessOption": matched_option,
            "status": "configured_before_submit",
        }

        save_json(
            "summary.json",
            summary,
        )

        print()
        print("========================================")
        print("PRUACCESS CONFIGURATION COMPLETE")
        print("========================================")

        print(
            "Excel B2:",
            excel_name,
        )

        print(
            "Matched PruAccess:",
            matched_option["text"],
        )

        print(
            "PruAccess ID:",
            matched_option["value"],
        )

        print(
            "Start date:",
            start_date,
        )

        print(
            "End date:",
            end_date,
        )

        print(
            "View type:",
            selected_values["viewType"],
        )

        print(
            "Price type:",
            selected_values["fundPriceType"],
        )

        print()
        print(
            "The form has NOT been submitted yet."
        )

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
