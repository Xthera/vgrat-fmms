#!/usr/bin/env python3

"""
VGrat FMS - AI Fund Exposure Research

Purpose
-------
Research PRULink funds and extract:

    - Up to 3 geographic exposures
    - Up to 3 sector exposures

Names only. No percentages are stored.

Source priority
---------------
1. Prudential URL from Funds Links.xlsx Column A
2. Prudential / official fund documents discovered from web search
3. Official underlying fund-manager documents
4. Reputable secondary sources only when primary sources cannot be found

AI
--
Local Hugging Face model:

    Qwen/Qwen3-1.7B

The model does NOT browse the web itself.

Python retrieves candidate webpages/PDFs first, then the
local model analyses the retrieved source material.

Output
------
data/fund_exposure.json

Existing BID / holdings pipelines are not modified.
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pypdf import PdfReader
from transformers import AutoModelForCausalLM, AutoTokenizer


# ============================================================
# CONFIGURATION
# ============================================================

MODEL_NAME = os.getenv(
    "HF_MODEL",
    "Qwen/Qwen3-1.7B",
)

EXCEL_PATH = Path(
    os.getenv(
        "FUNDS_LINKS_FILE",
        "Funds Links.xlsx",
    )
)

OUTPUT_PATH = Path(
    os.getenv(
        "EXPOSURE_OUTPUT",
        "data/fund_exposure.json",
    )
)

MAX_SEARCH_RESULTS = 8
MAX_DOCUMENTS_PER_FUND = 5

MAX_PAGE_CHARS = 18000
MAX_PDF_CHARS = 24000

REQUEST_TIMEOUT = 30

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/154.0 Safari/537.36 "
    "VGrat-FMS-Fund-Exposure-Research"
)

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": USER_AGENT,
        "Accept-Language": "en-SG,en;q=0.9",
    }
)


# ============================================================
# FUND LOADING
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def load_funds_from_excel(path: Path) -> list[dict[str, str]]:
    """
    Column A:
        Prudential URL / controlling universe

    Column B:
        Exact PruAccess / Prudential Fund Name reference
    """

    if not path.exists():
        raise FileNotFoundError(
            f"Funds Links file not found: {path}"
        )

    workbook = load_workbook(
        filename=path,
        read_only=True,
        data_only=True,
    )

    sheet = workbook.active

    funds: list[dict[str, str]] = []

    for row_number, row in enumerate(
        sheet.iter_rows(
            min_row=2,
            values_only=True,
        ),
        start=2,
    ):
        url = clean_text(
            row[0] if len(row) > 0 else ""
        )

        fund_name = clean_text(
            row[1] if len(row) > 1 else ""
        )

        if not url and not fund_name:
            continue

        if not fund_name:
            fund_name = url

        funds.append(
            {
                "row": str(row_number),
                "fundName": fund_name,
                "prudentialUrl": url,
            }
        )

    workbook.close()

    return funds


# ============================================================
# HTTP
# ============================================================

def fetch_url(
    url: str,
) -> tuple[str, bytes] | None:

    try:
        response = SESSION.get(
            url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )

        if response.status_code >= 400:
            return None

        content_type = (
            response.headers
            .get("content-type", "")
            .lower()
        )

        return content_type, response.content

    except Exception as exc:
        print(
            f"      fetch failed: {url} -> {exc}"
        )
        return None


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_text(
    content: bytes,
) -> str:

    try:
        reader = PdfReader(
            io.BytesIO(content)
        )

        pages: list[str] = []

        for page in reader.pages:
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""

            if text:
                pages.append(text)

            if sum(len(x) for x in pages) >= MAX_PDF_CHARS:
                break

        text = "\n".join(pages)

        return text[:MAX_PDF_CHARS]

    except Exception:
        return ""


# ============================================================
# HTML EXTRACTION
# ============================================================

def extract_html_text(
    content: bytes,
) -> str:

    try:
        soup = BeautifulSoup(
            content,
            "html.parser",
        )

        for tag in soup(
            [
                "script",
                "style",
                "noscript",
                "svg",
                "nav",
                "footer",
            ]
        ):
            tag.decompose()

        text = soup.get_text(
            "\n",
            strip=True,
        )

        return text[:MAX_PAGE_CHARS]

    except Exception:
        return ""


# ============================================================
# DOCUMENT NORMALISATION
# ============================================================

def extract_document(
    url: str,
) -> str:

    result = fetch_url(url)

    if not result:
        return ""

    content_type, content = result

    if (
        "pdf" in content_type
        or url.lower().split("?")[0].endswith(".pdf")
    ):
        return extract_pdf_text(content)

    return extract_html_text(content)


# ============================================================
# SEARCH ENGINE
# ============================================================

def ddg_search(
    query: str,
    max_results: int = MAX_SEARCH_RESULTS,
) -> list[dict[str, str]]:

    """
    Uses DuckDuckGo HTML search directly.

    No API key required.
    """

    search_url = (
        "https://html.duckduckgo.com/html/?q="
        + quote_plus(query)
    )

    try:
        response = SESSION.get(
            search_url,
            timeout=REQUEST_TIMEOUT,
        )

        if response.status_code >= 400:
            return []

        soup = BeautifulSoup(
            response.text,
            "html.parser",
        )

        results: list[dict[str, str]] = []

        for result in soup.select(
            ".result"
        ):

            link = result.select_one(
                ".result__a"
            )

            if not link:
                continue

            href = link.get("href")

            if not href:
                continue

            title = clean_text(
                link.get_text(" ", strip=True)
            )

            snippet_element = result.select_one(
                ".result__snippet"
            )

            snippet = (
                clean_text(
                    snippet_element.get_text(
                        " ",
                        strip=True,
                    )
                )
                if snippet_element
                else ""
            )

            results.append(
                {
                    "title": title,
                    "url": href,
                    "snippet": snippet,
                }
            )

            if len(results) >= max_results:
                break

        return results

    except Exception as exc:
        print(
            f"      search failed: {exc}"
        )
        return []


# ============================================================
# URL FILTERING
# ============================================================

def is_useful_document_url(
    url: str,
) -> bool:

    lowered = url.lower()

    blocked = [
        "facebook.com",
        "instagram.com",
        "youtube.com",
        "linkedin.com",
        "x.com",
        "twitter.com",
        "reddit.com",
        "tiktok.com",
    ]

    if any(
        domain in lowered
        for domain in blocked
    ):
        return False

    return True


def source_priority(
    url: str,
) -> int:

    host = urlparse(url).netloc.lower()

    if "prudential.com.sg" in host:
        return 1

    known_managers = [
        "eastspring.com",
        "abrdn.com",
        "fullertonfund.com",
        "hsbc.com",
        "blackrock.com",
        "schroders.com",
        "fidelity.com",
        "jpmorgan.com",
        "jpmorganassetmanagement.com",
        "amundi.com",
        "invesco.com",
        "franklintempleton.com",
        "manulifeim.com",
        "pimco.com",
    ]

    if any(
        manager in host
        for manager in known_managers
    ):
        return 2

    return 3


# ============================================================
# RESEARCH
# ============================================================

def build_search_queries(
    fund_name: str,
) -> list[str]:

    return [
        f'"{fund_name}" sector allocation geography',
        f'"{fund_name}" geographic allocation sector allocation',
        f'"{fund_name}" factsheet geography sector',
        f'"{fund_name}" underlying fund factsheet',
        f'"{fund_name}" portfolio sector country',
    ]


def research_fund(
    fund: dict[str, str],
) -> list[dict[str, Any]]:

    fund_name = fund["fundName"]
    prudential_url = fund["prudentialUrl"]

    print(
        f"      researching: {fund_name}"
    )

    candidates: list[dict[str, Any]] = []

    # --------------------------------------------------------
    # 1. Controlling Prudential URL
    # --------------------------------------------------------

    if prudential_url:
        candidates.append(
            {
                "url": prudential_url,
                "title": "Prudential source from Funds Links.xlsx",
                "snippet": "",
                "priority": 1,
            }
        )

    # --------------------------------------------------------
    # 2. Search web
    # --------------------------------------------------------

    seen_urls = {
        prudential_url
    } if prudential_url else set()

    for query in build_search_queries(
        fund_name
    ):

        print(
            f"        search: {query}"
        )

        results = ddg_search(query)

        for result in results:

            url = result["url"]

            if not is_useful_document_url(url):
                continue

            if url in seen_urls:
                continue

            seen_urls.add(url)

            candidates.append(
                {
                    **result,
                    "priority": source_priority(url),
                }
            )

            if len(candidates) >= MAX_DOCUMENTS_PER_FUND:
                break

        if len(candidates) >= MAX_DOCUMENTS_PER_FUND:
            break

        time.sleep(0.5)

    # --------------------------------------------------------
    # 3. Priority sort
    # --------------------------------------------------------

    candidates.sort(
        key=lambda item: item.get(
            "priority",
            99,
        )
    )

    documents: list[dict[str, Any]] = []

    for candidate in candidates[
        :MAX_DOCUMENTS_PER_FUND
    ]:

        url = candidate["url"]

        print(
            f"        document: {url}"
        )

        text = extract_document(url)

        if not text:
            continue

        documents.append(
            {
                "url": url,
                "title": candidate.get(
                    "title",
                    "",
                ),
                "snippet": candidate.get(
                    "snippet",
                    "",
                ),
                "text": text,
                "priority": candidate.get(
                    "priority",
                    99,
                ),
            }
        )

    return documents


# ============================================================
# AI
# ============================================================

def load_ai_model():

    print(
        f"Loading local Hugging Face model: {MODEL_NAME}"
    )

    tokenizer = AutoTokenizer.from_pretrained(
        MODEL_NAME,
        trust_remote_code=True,
    )

    model = AutoModelForCausalLM.from_pretrained(
        MODEL_NAME,
        torch_dtype="auto",
        device_map="auto",
        trust_remote_code=True,
    )

    return tokenizer, model


def build_ai_prompt(
    fund_name: str,
    documents: list[dict[str, Any]],
) -> str:

    source_text_parts: list[str] = []

    for index, document in enumerate(
        documents,
        start=1,
    ):

        source_text_parts.append(
            f"""
SOURCE {index}

URL:
{document["url"]}

TITLE:
{document["title"]}

CONTENT:
{document["text"]}
""".strip()
        )

    source_text = "\n\n".join(
        source_text_parts
    )

    return f"""
You are researching one investment fund.

Fund:
{fund_name}

Your task is ONLY to identify:

1. The top geographic exposures.
2. The top sector exposures.

IMPORTANT RULES:

- Use ONLY information explicitly supported by the supplied sources.
- Do NOT use your general knowledge.
- Do NOT guess.
- Do NOT infer geography from the fund name.
- Do NOT infer sector from the fund name.
- Do NOT invent missing categories.
- Return a maximum of 3 geography names.
- Return a maximum of 3 sector names.
- If only 1 valid geography is available, return 1.
- If only 2 valid geographies are available, return 2.
- If no reliable geography is available, return [].
- The same rule applies to sectors.
- Return names only.
- Do NOT return percentages.
- Do NOT return weights.
- Do NOT return explanations inside the arrays.
- Use the source's actual category names where possible.
- Rank the categories from highest exposure to lowest exposure when the source provides enough information to determine ranking.
- If the source explicitly identifies a "Top 3" list, use that.
- If the source contains a complete allocation table, determine the largest exposures from that table.
- Do not treat individual company names as sectors.
- Do not treat individual company names as geographies.
- "United States" and "USA" should be normalised to "United States".
- "UK" and "United Kingdom" should be normalised to "United Kingdom".
- "Hong Kong SAR" may be normalised to "Hong Kong".
- Preserve meaningful sector categories such as Technology, Financials, Healthcare, Industrials, Energy, Real Estate, Consumer Staples, Consumer Discretionary, Communication Services and Utilities.

Return ONLY valid JSON in exactly this structure:

{{
  "geography": [],
  "sector": []
}}

No markdown.
No commentary.
No percentages.

SOURCES:

{source_text}
""".strip()


def extract_json_object(
    text: str,
) -> dict[str, Any] | None:

    text = text.strip()

    # Remove accidental markdown fences.
    text = re.sub(
        r"^```(?:json)?",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"```$",
        "",
        text,
    )

    text = text.strip()

    # First attempt: complete JSON.
    try:
        parsed = json.loads(text)

        if isinstance(parsed, dict):
            return parsed

    except Exception:
        pass

    # Second attempt: locate JSON object.
    match = re.search(
        r"\{.*\}",
        text,
        flags=re.DOTALL,
    )

    if not match:
        return None

    try:
        parsed = json.loads(
            match.group(0)
        )

        if isinstance(parsed, dict):
            return parsed

    except Exception:
        return None

    return None


# ============================================================
# RESULT VALIDATION
# ============================================================

def clean_exposure_names(
    values: Any,
) -> list[str]:

    if not isinstance(values, list):
        return []

    cleaned: list[str] = []

    for value in values:

        if not isinstance(
            value,
            str,
        ):
            continue

        value = clean_text(value)

        if not value:
            continue

        # Remove accidental percentage text.
        value = re.sub(
            r"\s*\(?\d+(?:\.\d+)?%\)?",
            "",
            value,
        ).strip()

        if not value:
            continue

        if value.lower() in {
            "other",
            "others",
            "n/a",
            "na",
            "unknown",
            "not available",
            "unavailable",
            "none",
        }:
            continue

        if value not in cleaned:
            cleaned.append(value)

        if len(cleaned) >= 3:
            break

    return cleaned


def validate_ai_result(
    result: dict[str, Any] | None,
) -> dict[str, list[str]]:

    if not result:
        return {
            "geography": [],
            "sector": [],
        }

    return {
        "geography": clean_exposure_names(
            result.get(
                "geography",
                [],
            )
        ),
        "sector": clean_exposure_names(
            result.get(
                "sector",
                [],
            )
        ),
    }


# ============================================================
# PROCESS ONE FUND
# ============================================================

def process_fund(
    fund: dict[str, str],
    tokenizer,
    model,
) -> dict[str, Any]:

    fund_name = fund["fundName"]

    documents = research_fund(
        fund
    )

    if not documents:

        print(
            "      no usable source documents"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": fund[
                "prudentialUrl"
            ],
            "geography": [],
            "sector": [],
            "status": "no_source",
            "sources": [],
        }

    prompt = build_ai_prompt(
        fund_name,
        documents,
    )

    messages = [
        {
            "role": "system",
            "content": (
                "You extract structured financial "
                "information from supplied source documents. "
                "Never invent information."
            ),
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    print(
        "      running local AI..."
    )

    try:

        inputs = tokenizer.apply_chat_template(
            messages,
            tokenize=True,
            add_generation_prompt=True,
            return_tensors="pt",
        )

        inputs = inputs.to(
            model.device
        )

        generated = model.generate(
            inputs,
            max_new_tokens=180,
            do_sample=False,
            temperature=None,
            top_p=None,
        )

        generated_tokens = generated[
            0
        ][
            inputs.shape[-1]:
        ]

        response = tokenizer.decode(
            generated_tokens,
            skip_special_tokens=True,
        )

        parsed = extract_json_object(
            response
        )

        exposures = validate_ai_result(
            parsed
        )

        if exposures["geography"] or exposures["sector"]:

            status = "success"

        else:

            status = "no_exposure_found"

        return {
            "fundName": fund_name,
            "prudentialUrl": fund[
                "prudentialUrl"
            ],
            "geography": exposures[
                "geography"
            ],
            "sector": exposures[
                "sector"
            ],
            "status": status,
            "sources": [
                {
                    "title": document[
                        "title"
                    ],
                    "url": document[
                        "url"
                    ],
                }
                for document in documents
            ],
        }

    except Exception as exc:

        print(
            f"      AI failed: {exc}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": fund[
                "prudentialUrl"
            ],
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": str(exc),
            "sources": [
                {
                    "title": document[
                        "title"
                    ],
                    "url": document[
                        "url"
                    ],
                }
                for document in documents
            ],
        }


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    parser = argparse.ArgumentParser(
        description=(
            "Research PRULink fund geography "
            "and sector exposures using local AI."
        )
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help=(
            "Process only the first N funds. "
            "0 = all funds."
        ),
    )

    args = parser.parse_args()

    print("=" * 70)
    print("VGrat FMS - AI Fund Exposure Research")
    print("=" * 70)

    print(
        f"Excel: {EXCEL_PATH}"
    )

    print(
        f"Model: {MODEL_NAME}"
    )

    print(
        f"Output: {OUTPUT_PATH}"
    )

    # --------------------------------------------------------
    # Load funds
    # --------------------------------------------------------

    funds = load_funds_from_excel(
        EXCEL_PATH
    )

    if args.limit > 0:
        funds = funds[
            :args.limit
        ]

    print(
        f"Funds loaded: {len(funds)}"
    )

    # --------------------------------------------------------
    # Load model once
    # --------------------------------------------------------

    tokenizer, model = load_ai_model()

    # --------------------------------------------------------
    # Process funds
    # --------------------------------------------------------

    results: list[dict[str, Any]] = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        print()
        print(
            f"[{index}/{len(funds)}] "
            f"{fund['fundName']}"
        )

        result = process_fund(
            fund,
            tokenizer,
            model,
        )

        results.append(
            result
        )

        print(
            f"      geography: "
            f"{result.get('geography', [])}"
        )

        print(
            f"      sector: "
            f"{result.get('sector', [])}"
        )

        print(
            f"      status: "
            f"{result.get('status')}"
        )

    # --------------------------------------------------------
    # Build output
    # --------------------------------------------------------

    successful = sum(
        1
        for result in results
        if result.get("status")
        == "success"
    )

    output = {
        "generatedAtUTC": time.strftime(
            "%Y-%m-%dT%H:%M:%SZ",
            time.gmtime(),
        ),
        "model": MODEL_NAME,
        "rules": {
            "maxGeography": 3,
            "maxSector": 3,
            "percentagesStored": False,
            "inferredValuesAllowed": False,
        },
        "summary": {
            "fundsProcessed": len(results),
            "successful": successful,
            "withoutReliableExposure": (
                len(results) - successful
            ),
        },
        "funds": results,
    }

    OUTPUT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with OUTPUT_PATH.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            output,
            file,
            indent=2,
            ensure_ascii=False,
        )

    print()
    print("=" * 70)
    print("COMPLETE")
    print("=" * 70)

    print(
        f"Processed: {len(results)}"
    )

    print(
        f"Successful: {successful}"
    )

    print(
        f"Output: {OUTPUT_PATH}"
    )

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
