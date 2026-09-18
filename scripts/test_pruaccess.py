import asyncio
import json
from pathlib import Path

from playwright.async_api import async_playwright


TEST_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

TEST_FUND_NAME = (
    "PRULink ActiveInvest Portfolio - Balanced (SGD)"
)

TEST_FUND_ID = "335771"

OUTPUT_DIR = Path("output_pruaccess")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


async def main():

    captured_requests = []
    captured_responses = []

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        page = await browser.new_page(
            viewport={
                "width": 1440,
                "height": 1200
            }
        )

        async def handle_request(request):

            captured_requests.append({
                "method": request.method,
                "url": request.url,
                "resourceType": request.resource_type,
                "postData": request.post_data,
                "headers": dict(request.headers),
            })

            print()
            print("REQUEST:")
            print(request.method, request.url)

            if request.post_data:
                print("POST DATA:")
                print(request.post_data)

        async def handle_response(response):

            captured_responses.append({
                "url": response.url,
                "status": response.status,
                "contentType": response.headers.get(
                    "content-type",
                    ""
                ),
            })

            print()
            print("RESPONSE:")
            print(response.status, response.url)

            content_type = response.headers.get(
                "content-type",
                ""
            ).lower()

            interesting = any(
                keyword in response.url.lower()
                for keyword in [
                    "fund",
                    "performance",
                    "price",
                    "history",
                    "ajax",
                    "json",
                    "api",
                    "switch",
                ]
            )

            if interesting or "text/html" in content_type:

                try:

                    body = await response.text()

                    filename = (
                        f"response_"
                        f"{len(captured_responses)}.txt"
                    )

                    output_file = (
                        OUTPUT_DIR / filename
                    )

                    output_file.write_text(
                        body,
                        encoding="utf-8"
                    )

                    print(
                        "SAVED RESPONSE:",
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

        print("========================================")
        print("OPENING PRUACCESS")
        print("========================================")
        print(TEST_URL)

        try:

            await page.goto(
                TEST_URL,
                wait_until="domcontentloaded",
                timeout=120000
            )

        except Exception as e:

            print(
                "PAGE LOAD ERROR:",
                repr(e)
            )

        print()
        print("Waiting for initial JavaScript...")

        await page.wait_for_timeout(
            5000
        )

        # --------------------------------------------------
        # Inspect forms and buttons
        # --------------------------------------------------

        form_info = await page.evaluate(
            """
            () => {

                return {

                    forms: [...document.forms].map(
                        (form, index) => ({
                            index: index,
                            id: form.id,
                            name: form.name,
                            method: form.method,
                            action: form.action
                        })
                    ),

                    buttons: [
                        ...document.querySelectorAll(
                            'button, input[type="submit"], '
                            'input[type="button"]'
                        )
                    ].map(
                        (button, index) => ({
                            index: index,
                            tag: button.tagName,
                            id: button.id,
                            name: button.name,
                            type: button.type,
                            value: button.value,
                            text: button.innerText || "",
                            outerHTML: button.outerHTML
                        })
                    )

                };

            }
            """
        )

        (
            OUTPUT_DIR / "forms_and_buttons.json"
        ).write_text(
            json.dumps(
                form_info,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("========================================")
        print("FORMS AND BUTTONS")
        print("========================================")

        print(
            json.dumps(
                form_info,
                indent=2,
                ensure_ascii=False
            )
        )

        # --------------------------------------------------
        # Select the test fund
        # --------------------------------------------------

        print()
        print("========================================")
        print("SELECTING TEST FUND")
        print("========================================")

        fund_select = page.locator(
            "#fundName"
        )

        await fund_select.select_option(
            TEST_FUND_ID
        )

        print(
            "Selected:",
            TEST_FUND_NAME
        )

        print(
            "Fund ID:",
            TEST_FUND_ID
        )

        # --------------------------------------------------
        # Select Bid Price
        # --------------------------------------------------

        price_select = page.locator(
            "#fundPriceType"
        )

        await price_select.select_option(
            "BID"
        )

        print(
            "Price type: BID"
        )

        # --------------------------------------------------
        # Set date range
        #
        # Use a long range. If PruAccess limits the
        # maximum range, the resulting page/request will
        # tell us what it accepts.
        # --------------------------------------------------

        start_date = page.locator(
            "#startDate"
        )

        end_date = page.locator(
            "#endDate"
        )

        await start_date.fill(
            "03-Nov-2021"
        )

        await end_date.fill(
            "17-Sep-2026"
        )

        print(
            "Start date:",
            await start_date.input_value()
        )

        print(
            "End date:",
            await end_date.input_value()
        )

        # --------------------------------------------------
        # Capture selected values before submission
        # --------------------------------------------------

        selected_values = await page.evaluate(
            """
            () => {

                const fund = document.querySelector(
                    '#fundName'
                );

                const price = document.querySelector(
                    '#fundPriceType'
                );

                const start = document.querySelector(
                    '#startDate'
                );

                const end = document.querySelector(
                    '#endDate'
                );

                return {
                    fundValue: fund ? fund.value : null,
                    fundText: fund
                        ? fund.options[fund.selectedIndex].text
                        : null,
                    priceType: price
                        ? price.value
                        : null,
                    startDate: start
                        ? start.value
                        : null,
                    endDate: end
                        ? end.value
                        : null
                };

            }
            """
        )

        (
            OUTPUT_DIR / "selected_values.json"
        ).write_text(
            json.dumps(
                selected_values,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("SELECTED VALUES:")
        print(
            json.dumps(
                selected_values,
                indent=2,
                ensure_ascii=False
            )
        )

        # --------------------------------------------------
        # Attempt to submit the form
        # --------------------------------------------------

        print()
        print("========================================")
        print("SUBMITTING PRUACCESS FORM")
        print("========================================")

        submitted = False

        # First look for common submit controls.
        submit_candidates = page.locator(
            'button[type="submit"], '
            'input[type="submit"], '
            'button'
        )

        count = await submit_candidates.count()

        print(
            "Submit candidates:",
            count
        )

        for i in range(count):

            try:

                element = submit_candidates.nth(i)

                tag = await element.evaluate(
                    "(el) => el.tagName"
                )

                element_id = await element.get_attribute(
                    "id"
                )

                name = await element.get_attribute(
                    "name"
                )

                value = await element.get_attribute(
                    "value"
                )

                text = (
                    await element.inner_text()
                ).strip()

                print(
                    f"Candidate {i}: "
                    f"tag={tag}, "
                    f"id={element_id}, "
                    f"name={name}, "
                    f"value={value}, "
                    f"text={text}"
                )

            except Exception as e:

                print(
                    "Could not inspect candidate:",
                    i,
                    repr(e)
                )

        # Try buttons that look like a search/query/
        # performance submission control.
        for i in range(count):

            element = submit_candidates.nth(i)

            try:

                label = " ".join([
                    str(
                        await element.get_attribute("id")
                        or ""
                    ),
                    str(
                        await element.get_attribute("name")
                        or ""
                    ),
                    str(
                        await element.get_attribute("value")
                        or ""
                    ),
                    (
                        await element.inner_text()
                    ).strip()
                ]).lower()

                keywords = [
                    "search",
                    "submit",
                    "view",
                    "performance",
                    "show",
                    "generate",
                    "go"
                ]

                if not any(
                    keyword in label
                    for keyword in keywords
                ):
                    continue

                print()
                print(
                    "Attempting click:",
                    label
                )

                try:

                    await element.click(
                        timeout=10000
                    )

                    submitted = True

                    print(
                        "CLICKED."
                    )

                    break

                except Exception as e:

                    print(
                        "Click failed:",
                        repr(e)
                    )

            except Exception:
                continue

        # --------------------------------------------------
        # Wait for resulting request/page
        # --------------------------------------------------

        if submitted:

            print()
            print(
                "Waiting for PruAccess result..."
            )

            await page.wait_for_timeout(
                10000
            )

        else:

            print()
            print(
                "No obvious submit button was "
                "automatically clicked."
            )

        # --------------------------------------------------
        # Save resulting page
        # --------------------------------------------------

        html = await page.content()

        (
            OUTPUT_DIR / "result.html"
        ).write_text(
            html,
            encoding="utf-8"
        )

        try:

            visible_text = await page.locator(
                "body"
            ).inner_text()

            (
                OUTPUT_DIR / "result_text.txt"
            ).write_text(
                visible_text,
                encoding="utf-8"
            )

        except Exception as e:

            print(
                "Could not capture result text:",
                repr(e)
            )

        # --------------------------------------------------
        # Save all network information
        # --------------------------------------------------

        (
            OUTPUT_DIR / "requests.json"
        ).write_text(
            json.dumps(
                captured_requests,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        (
            OUTPUT_DIR / "responses.json"
        ).write_text(
            json.dumps(
                captured_responses,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        summary = {
            "testUrl": TEST_URL,
            "testFundName": TEST_FUND_NAME,
            "testFundId": TEST_FUND_ID,
            "selectedValues": selected_values,
            "submitted": submitted,
            "requestCount": len(
                captured_requests
            ),
            "responseCount": len(
                captured_responses
            )
        }

        (
            OUTPUT_DIR / "summary.json"
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
        print("PRUACCESS TEST COMPLETE")
        print("========================================")

        print(
            "Submitted:",
            submitted
        )

        print(
            "Requests:",
            len(captured_requests)
        )

        print(
            "Responses:",
            len(captured_responses)
        )

        print()
        print(
            "Files saved to:",
            OUTPUT_DIR
        )

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
