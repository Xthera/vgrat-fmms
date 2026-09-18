import asyncio
import json
from pathlib import Path

from playwright.async_api import async_playwright


TEST_URL = "https://pruaccess.prudential.com.sg/prulinkfund/viewFundPerformance.do"

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
            viewport={"width": 1440, "height": 1000}
        )

        async def handle_request(request):

            url = request.url

            captured_requests.append({
                "method": request.method,
                "url": url,
                "resourceType": request.resource_type,
                "postData": request.post_data,
                "headers": dict(request.headers),
            })

            print()
            print("REQUEST:")
            print(request.method, url)

            if request.post_data:
                print("POST DATA:")
                print(request.post_data)

        async def handle_response(response):

            url = response.url

            captured_responses.append({
                "url": url,
                "status": response.status,
                "contentType": response.headers.get(
                    "content-type",
                    ""
                ),
            })

            print()
            print("RESPONSE:")
            print(response.status, url)

            content_type = response.headers.get(
                "content-type",
                ""
            ).lower()

            # Save potentially useful responses.
            interesting = any(
                keyword in url.lower()
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

            if interesting:

                try:

                    body = await response.text()

                    filename = (
                        f"response_{len(captured_responses)}.txt"
                    )

                    output_file = OUTPUT_DIR / filename

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
        print("Waiting for page JavaScript...")

        await page.wait_for_timeout(
            15000
        )

        # Capture rendered page.
        html = await page.content()

        (
            OUTPUT_DIR / "rendered.html"
        ).write_text(
            html,
            encoding="utf-8"
        )

        # Capture visible text.
        try:

            visible_text = await page.locator(
                "body"
            ).inner_text()

            (
                OUTPUT_DIR / "visible_text.txt"
            ).write_text(
                visible_text,
                encoding="utf-8"
            )

        except Exception as e:

            print(
                "Could not capture visible text:",
                repr(e)
            )

        # Capture page forms and select elements.
        try:

            form_data = await page.evaluate(
                """
                () => {

                    const selects = [...document.querySelectorAll("select")];

                    return {
                        selects: selects.map((s, index) => ({
                            index: index,
                            name: s.name,
                            id: s.id,
                            value: s.value,
                            options: [...s.options].map(o => ({
                                text: o.text,
                                value: o.value
                            }))
                        })),

                        inputs: [...document.querySelectorAll("input")].map((i, index) => ({
                            index: index,
                            name: i.name,
                            id: i.id,
                            type: i.type,
                            value: i.value
                        }))
                    };

                }
                """
            )

            (
                OUTPUT_DIR / "form_elements.json"
            ).write_text(
                json.dumps(
                    form_data,
                    indent=2,
                    ensure_ascii=False
                ),
                encoding="utf-8"
            )

        except Exception as e:

            print(
                "Could not inspect forms:",
                repr(e)
            )

        # Save request log.
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

        # Save response log.
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

        print()
        print("========================================")
        print("PRUACCESS DIAGNOSTIC COMPLETE")
        print("========================================")

        print(
            "Requests captured:",
            len(captured_requests)
        )

        print(
            "Responses captured:",
            len(captured_responses)
        )

        print()
        print("Files saved to:")
        print(OUTPUT_DIR)

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
