import asyncio
import json
from pathlib import Path
from playwright.async_api import async_playwright


TEST_URL = (
    "https://www.prudential.com.sg/en/products/wealth/ilp/"
    "prulink-funds/"
    "prulink-activeinvest-portfolio-balanced-accumulation/"
)

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


IMPORTANT_PARTS = [
    "ilpfunds.json",
    "ilpseries.json",
]


async def main():

    captured_json = []
    captured_requests = []

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        page = await browser.new_page(
            viewport={"width": 1440, "height": 1000}
        )

        async def handle_response(response):

            url = response.url

            if not any(part in url for part in IMPORTANT_PARTS):
                return

            print()
            print("FOUND PRUDENTIAL JSON:")
            print(url)
            print("STATUS:", response.status)

            try:
                body = await response.text()

                filename = f"response_{len(captured_json) + 1}.json"

                output_file = OUTPUT_DIR / filename

                output_file.write_text(
                    body,
                    encoding="utf-8"
                )

                captured_json.append({
                    "filename": filename,
                    "url": url,
                    "status": response.status,
                    "contentType": response.headers.get(
                        "content-type",
                        ""
                    )
                })

            except Exception as e:

                print(
                    "Could not read response:",
                    repr(e)
                )

        async def handle_request(request):

            url = request.url

            if any(part in url for part in IMPORTANT_PARTS):

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

        print("Opening Prudential page...")
        print(TEST_URL)

        await page.goto(
            TEST_URL,
            wait_until="domcontentloaded",
            timeout=120000
        )

        print("Waiting for Prudential JavaScript...")

        await page.wait_for_timeout(
            15000
        )

        # Scroll through the page so lazy-loaded
        # content has an opportunity to load.
        await page.evaluate(
            """
            window.scrollTo(0, document.body.scrollHeight);
            """
        )

        await page.wait_for_timeout(
            5000
        )

        # Save the rendered page for reference.
        html = await page.content()

        (
            OUTPUT_DIR / "rendered.html"
        ).write_text(
            html,
            encoding="utf-8"
        )

        # Save visible page text.
        visible_text = await page.locator(
            "body"
        ).inner_text()

        (
            OUTPUT_DIR / "visible_text.txt"
        ).write_text(
            visible_text,
            encoding="utf-8"
        )

        # Save captured request list.
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

        # Save summary.
        summary = {
            "testUrl": TEST_URL,
            "jsonResponsesCaptured": len(
                captured_json
            ),
            "responses": captured_json
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
        print("PRUDENTIAL JSON CAPTURE COMPLETE")
        print("========================================")
        print(
            "JSON responses:",
            len(captured_json)
        )

        for item in captured_json:
            print(
                item["filename"],
                "->",
                item["url"]
            )

        print()
        print("Files saved to output/")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
