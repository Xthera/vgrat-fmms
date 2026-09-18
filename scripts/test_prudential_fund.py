import asyncio
import json
import re
from pathlib import Path
from playwright.async_api import async_playwright


TEST_URL = (
    "https://www.prudential.com.sg/content/prudential-aem-lbu/pacs/en/"
    "products/wealth/ilp/prulink-funds/"
    "prulink-activeinvest-portfolio-balanced-accumulation.html"
)

OUTPUT_DIR = Path("output")
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
            url = request.url.lower()

            keywords = [
                "api",
                "fund",
                "fundinfo",
                "performance",
                "price",
                "bid",
                "offer",
                "dividend",
                "factsheet",
                "citicode",
                "d39f",
            ]

            if any(keyword in url for keyword in keywords):
                captured_requests.append({
                    "method": request.method,
                    "url": request.url,
                    "resourceType": request.resource_type,
                    "postData": request.post_data,
                })

        async def handle_response(response):
            url = response.url.lower()

            keywords = [
                "api",
                "fund",
                "fundinfo",
                "performance",
                "price",
                "bid",
                "offer",
                "dividend",
                "factsheet",
                "citicode",
                "d39f",
            ]

            if any(keyword in url for keyword in keywords):
                captured_responses.append({
                    "status": response.status,
                    "url": response.url,
                    "contentType": response.headers.get(
                        "content-type", ""
                    ),
                })

        page.on("request", handle_request)
        page.on("response", handle_response)

        print("Opening Prudential fund page...")
        await page.goto(
            TEST_URL,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        # Give the page's JavaScript time to load its data.
        await page.wait_for_timeout(15000)

        # Try to trigger lazy-loaded content.
        await page.evaluate(
            """
            window.scrollTo(0, document.body.scrollHeight);
            """
        )

        await page.wait_for_timeout(5000)

        # Save complete rendered HTML.
        html = await page.content()
        (OUTPUT_DIR / "rendered.html").write_text(
            html,
            encoding="utf-8"
        )

        # Save visible text.
        visible_text = await page.locator("body").inner_text()
        (OUTPUT_DIR / "visible_text.txt").write_text(
            visible_text,
            encoding="utf-8"
        )

        # Save all links.
        links = await page.locator("a").evaluate_all(
            """
            elements => elements.map(a => ({
                text: (a.innerText || "").trim(),
                href: a.href
            }))
            """
        )

        (OUTPUT_DIR / "links.json").write_text(
            json.dumps(
                links,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # Find references to D39F.
        d39f_matches = []

        for match in re.finditer(
            r".{0,250}D39F.{0,500}",
            html,
            flags=re.IGNORECASE | re.DOTALL,
        ):
            d39f_matches.append(match.group(0))

        (OUTPUT_DIR / "d39f_matches.txt").write_text(
            "\n\n--- MATCH ---\n\n".join(d39f_matches),
            encoding="utf-8"
        )

        # Save captured requests.
        (OUTPUT_DIR / "requests.json").write_text(
            json.dumps(
                captured_requests,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # Save captured responses.
        (OUTPUT_DIR / "responses.json").write_text(
            json.dumps(
                captured_responses,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # Basic page information.
        title = await page.title()

        summary = {
            "testUrl": TEST_URL,
            "pageTitle": title,
            "capturedRequestCount": len(captured_requests),
            "capturedResponseCount": len(captured_responses),
            "d39fMatchCount": len(d39f_matches),
        }

        (OUTPUT_DIR / "summary.json").write_text(
            json.dumps(
                summary,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        print()
        print("========================================")
        print("PRUDENTIAL TEST COMPLETE")
        print("========================================")
        print(f"Page title: {title}")
        print(
            f"Captured requests: "
            f"{len(captured_requests)}"
        )
        print(
            f"Captured responses: "
            f"{len(captured_responses)}"
        )
        print(
            f"D39F matches: "
            f"{len(d39f_matches)}"
        )
        print()
        print("Diagnostic files written to:")
        print("output/")
        print()

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
