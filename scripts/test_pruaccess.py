
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
    Read the first fund from Funds Links.xlsm.

    Column A:
        Prudential fund URL

    Column B:
        PruAccess fund name

    B2 is used to select the PruAccess fund.
    A2 is used to obtain the Prudential fund information.
    """

    print()
    print("========================================")
    print("READING FUNDS_LINKS.XLSM")
    print("========================================")

    print("Expected file:")
    print(EXCEL_FILE)

    if not EXCEL_FILE.exists():
        print()
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

        print()
        print("Worksheet:")
        print(sheet.title)

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
    Normalise text for safe fund-name comparison.
    """

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

async def get_prudential_inception_date(
    page,
    prudential_url,
):
    """
    Open the Prudential fund page and capture:

        ilpfunds.json

    The Prudential inceptionDate is used as the
    PruAccess start date.
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

        print("STATUS:")
        print(response.status)

        try:
            body = await response.text()

            data = json.loads(body)

            captured["url"] = response_url
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

    if not isinstance(data, list):
        raise RuntimeError(
            "Unexpected Prudential ilpfunds.json structure."
        )

    if not data:
        raise RuntimeError(
            "Prudential ilpfunds.json returned no fund data."
        )

    fund = data[0]

    fund_name = fund.get(
        "fundName",
        "",
    )

    inception_date = fund.get(
        "inceptionDate",
        "",
    )

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


# ============================================================
# DATE CONVERSION
# ============================================================

def convert_prudential_date_to_pruaccess(
    value,
):
    """
    Convert:

        03/11/2021

    to:

        03-Nov-2021
    """

    parsed = datetime.strptime(
        value,
        "%d/%m/%Y",
    )

    return parsed.strftime(
        "%d-%b-%Y"
    )


def get_run_date():
    """
    Date on which this GitHub Actions request is made.
    """

    return datetime.now().strftime(
        "%d-%b-%Y"
    )


# ============================================================
# SAVE JSON
# ============================================================

def save_json(
    filename,
    data,
):
    output_file = OUTPUT_DIR / filename

    output_file.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print()
    print("Saved:")
    print(output_file)


# ============================================================
# MAIN
# ============================================================

async def main():

    # ========================================================
    # 1. READ EXCEL
    # ========================================================

    excel_fund = read_excel_fund()

    save_json(
        "excel_fund.json",
        excel_fund,
    )

    # ========================================================
    # 2. START BROWSER
    # ========================================================

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True,
        )

        page = await browser.new_page(
            viewport={
                "width": 1440,
                "height": 1000,
            }
        )

        # ====================================================
        # 3. GET PRUDENTIAL INCEPTION DATE
        # ====================================================

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

        # ====================================================
        # 4. CONVERT DATES
        #
```
