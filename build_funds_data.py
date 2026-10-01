#!/usr/bin/env python3

"""
VGrat FMS - BUILD FUNDS DATA

Merges the outputs of the extraction pipelines into two data files:

```
data/funds.json
    fund information
    online geographic research
    online sector research
    Top Holdings

data/bid_history.json
    historical BID observations
```

# MASTER UNIVERSE

Funds Links.xlsm, Column A (Prudential URL) / Column B (PruAccess name).

Every populated Excel row is evaluated independently.

# PER-FUND PUBLICATION RULE

1. SUCCESS + previous record exists
   -> replace the previous record with newly extracted data.

2. SUCCESS + no previous record
   -> insert the newly extracted fund.

3. FAILURE + previous record exists
   -> retain the previous record unchanged.

4. FAILURE + no previous record
   -> do NOT publish a record for that fund.

A failure for one fund NEVER prevents successful funds from being
published.

# ONLINE FUND RESEARCH

Research is performed as an enrichment step.

It does NOT determine whether the core fund record is successful.

Research source hierarchy:

```
1. Prudential online fund page
   - identify underlying fund
   - identify underlying investment manager

2. Underlying manager's online HTML fund page
   - country / geographic allocation
   - sector allocation

3. Authoritative secondary online HTML source where available
```

# PDF RESTRICTIONS

This research layer does NOT:

```
- download Prudential PDFs
- parse Prudential PDFs
- use pdfplumber
- crawl PDF allocation tables
```

PDF URLs are explicitly rejected by the online research crawler.

# GEOGRAPHIC EXPOSURE

Only the top 2 regions are published.

Regions are ranked using the sum of published country weights.

Once a region is selected, ALL published countries belonging to that
region are retained.

No country limit is imposed within a selected region.

Country weights are never invented.

Regional weights are never invented.

Countries that cannot be confidently mapped to a region are not used
to create a regional classification.

# SECTOR EXPOSURE

Only the top 5 published sectors are retained.

Sector weights must come from an online allocation source.

No sector is inferred from:

```
- fund name
- company name
- Top Holdings
- investment objective
```

# DIVIDEND NORMALIZATION

Final funds.json rule:

```
dividendRate non-empty -> hasDividend = true
dividendRate empty/missing -> hasDividend = false
```

The existing dividendRate and dividendUnits values are never modified.

# IMPORTANT

Previous published data is used ONLY as fallback for the core fund/BID
publication process.

Previous research data is NOT used as newly retrieved research data.

If online research fails, the current fund can still be successfully
published with:

```
research.status = "unresolved"
```

and empty research arrays.

# OUTPUT

data/funds.json

Each fund may contain:

```
geographicExposure
topSectors
research
```

Existing:

```
fund
topHoldings
dividendRate
dividendUnits
hasDividend
```

remain intact.
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, unquote, urljoin, urlparse
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

# =============================================================================

# CONFIGURATION

# =============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

BASELINE_FILE = Path("output_holdings/all_holdings.json")

RECOVERY_STAGES = [
("recovery1", Path("output_holdings_recovery")),
("recovery2", Path("output_holdings_recovery_2")),
("recovery3", Path("output_holdings_recovery_3")),
]

PRUACCESS_FUNDS_DIR = Path("output_pruaccess/funds")

DATA_DIR = Path("data")
FUNDS_OUT = DATA_DIR / "funds.json"
BID_OUT = DATA_DIR / "bid_history.json"

SCHEMA_VERSION = 1
MAX_HOLDINGS = 10

MAX_RESEARCH_REGIONS = 2
MAX_RESEARCH_SECTORS = 5

RESEARCH_TIMEOUT_SECONDS = 25
RESEARCH_DELAY_SECONDS = 0.35

SINGAPORE_TZ = ZoneInfo("Asia/Singapore")

USER_AGENT = (
"Mozilla/5.0 "
"(Windows NT 10.0; Win64; x64) "
"AppleWebKit/537.36 "
"(KHTML, like Gecko) "
"Chrome/154.0 Safari/537.36 "
"VGrat-FMS/1.0"
)

# Retained for compatibility with existing workflow/environment.

ALLOW_UNRESOLVED = os.environ.get(
"ALLOW_UNRESOLVED",
"",
).strip().lower() in {"1", "true", "yes"}

# =============================================================================

# REGION MAP

# =============================================================================

COUNTRY_REGIONS = {
# North America
"United States": "North America",
"USA": "North America",
"United States of America": "North America",
"Canada": "North America",
"Mexico": "North America",

# Latin America
"Brazil": "Latin America",
"Argentina": "Latin America",
"Chile": "Latin America",
"Colombia": "Latin America",
"Peru": "Latin America",
"Panama": "Latin America",
"Uruguay": "Latin America",
"Ecuador": "Latin America",

# Europe
"United Kingdom": "Europe",
"UK": "Europe",
"Germany": "Europe",
"France": "Europe",
"Italy": "Europe",
"Spain": "Europe",
"Netherlands": "Europe",
"Belgium": "Europe",
"Switzerland": "Europe",
"Sweden": "Europe",
"Denmark": "Europe",
"Norway": "Europe",
"Finland": "Europe",
"Ireland": "Europe",
"Austria": "Europe",
"Portugal": "Europe",
"Greece": "Europe",
"Poland": "Europe",
"Czech Republic": "Europe",
"Hungary": "Europe",
"Romania": "Europe",

# Asia Pacific
"Japan": "Asia Pacific",
"China": "Asia Pacific",
"Hong Kong": "Asia Pacific",
"Taiwan": "Asia Pacific",
"Taiwan (Republic of China)": "Asia Pacific",
"South Korea": "Asia Pacific",
"Korea": "Asia Pacific",
"Singapore": "Asia Pacific",
"Australia": "Asia Pacific",
"New Zealand": "Asia Pacific",
"India": "Asia Pacific",
"Indonesia": "Asia Pacific",
"Malaysia": "Asia Pacific",
"Thailand": "Asia Pacific",
"Philippines": "Asia Pacific",
"Vietnam": "Asia Pacific",

# Middle East & Africa
"Israel": "Middle East & Africa",
"Saudi Arabia": "Middle East & Africa",
"United Arab Emirates": "Middle East & Africa",
"Qatar": "Middle East & Africa",
"Kuwait": "Middle East & Africa",
"South Africa": "Middle East & Africa",
"Egypt": "Middle East & Africa",
"Morocco": "Middle East & Africa",
"Nigeria": "Middle East & Africa",

# Other Europe / Eurasia
"Russia": "Europe",
"Turkey": "Europe",
```

}

# =============================================================================

# TEXT / TIME HELPERS

# =============================================================================

def clean_text(value) -> str:
if value is None:
return ""

```
return re.sub(
    r"\s+",
    " ",
    str(value).replace("\xa0", " "),
).strip()
```

def utc_now_iso() -> str:
return (
datetime.now(timezone.utc)
.replace(microsecond=0)
.isoformat()
)

def singapore_today() -> str:
return datetime.now(
SINGAPORE_TZ
).date().isoformat()

def load_json(path: Path):
try:
return json.loads(
path.read_text(
encoding="utf-8"
)
)
except Exception as error:
raise RuntimeError(
f"Could not read {path}: {error}"
) from error

def write_json_atomic(
path: Path,
data,
compact: bool = False,
) -> None:

```
path.parent.mkdir(
    parents=True,
    exist_ok=True,
)

tmp = path.with_suffix(
    path.suffix + ".tmp"
)

if compact:
    text = json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )
else:
    text = json.dumps(
        data,
        indent=2,
        ensure_ascii=False,
    )

tmp.write_text(
    text + "\n",
    encoding="utf-8",
)

tmp.replace(path)
```

# =============================================================================

# PREVIOUS FILES

# =============================================================================

def load_previous_file(path: Path) -> dict:

```
if not path.exists():

    print(
        f"INFO: Previous output not found: {path}"
    )

    return {
        "schemaVersion": SCHEMA_VERSION,
        "funds": [],
    }

try:
    data = load_json(path)

except Exception as error:

    raise RuntimeError(
        f"Previous published file exists but cannot be read: "
        f"{path}: {error}"
    ) from error

if not isinstance(data, dict):

    raise RuntimeError(
        f"Previous published file has invalid envelope: {path}"
    )

if not isinstance(data.get("funds"), list):

    raise RuntimeError(
        f"Previous published file has no valid funds list: {path}"
    )

return data
```

def index_previous_records(
data: dict,
) -> dict[int, dict]:

```
result = {}

for item in data.get("funds", []):

    if not isinstance(item, dict):
        continue

    row = item.get("excelRow")

    if isinstance(row, bool):
        continue

    try:
        row = int(row)
    except (TypeError, ValueError):
        continue

    result[row] = item

return result
```

# =============================================================================

# DIVIDEND NORMALIZATION

# =============================================================================

def normalize_dividend_fields(
fund_info: dict,
) -> dict:

```
dividend_rate = clean_text(
    fund_info.get("dividendRate")
)

fund_info["hasDividend"] = bool(
    dividend_rate
)

return fund_info
```

# =============================================================================

# EXCEL MASTER UNIVERSE

# =============================================================================

def read_excel_funds() -> dict[int, dict]:

```
if not EXCEL_FILE.exists():

    raise FileNotFoundError(
        f"Excel file not found: {EXCEL_FILE}"
    )

workbook = load_workbook(
    EXCEL_FILE,
    read_only=True,
    keep_vba=True,
    data_only=True,
)

funds = {}

try:

    worksheet = workbook.active

    for row in range(
        2,
        worksheet.max_row + 1,
    ):

        url = clean_text(
            worksheet.cell(
                row=row,
                column=1,
            ).value
        )

        name = clean_text(
            worksheet.cell(
                row=row,
                column=2,
            ).value
        )

        if not url:
            continue

        funds[row] = {
            "excelRow": row,
            "prudentialUrl": url,
            "pruAccessName": name,
        }

finally:

    workbook.close()

if not funds:

    raise RuntimeError(
        "No populated URLs in Excel Column A."
    )

return funds
```

# =============================================================================

# HTTP / ONLINE RESEARCH HELPERS

# =============================================================================

def fetch_html(
url: str,
timeout: int = RESEARCH_TIMEOUT_SECONDS,
) -> tuple[str, str]:

```
parsed = urlparse(url)

if parsed.scheme not in {"http", "https"}:
    raise ValueError(
        f"Unsupported URL scheme: {url}"
    )

if parsed.path.lower().endswith(".pdf"):
    raise ValueError(
        f"PDF research source rejected: {url}"
    )

request = Request(
    url,
    headers={
        "User-Agent": USER_AGENT,
        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,*/*;q=0.8"
        ),
        "Accept-Language": "en-SG,en;q=0.9",
    },
)

with urlopen(
    request,
    timeout=timeout,
) as response:

    final_url = response.geturl()

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        )
        .lower()
    )

    if (
        "pdf" in content_type
        or final_url.lower().split("?")[0].endswith(".pdf")
    ):
        raise ValueError(
            f"PDF research source rejected: {final_url}"
        )

    raw = response.read()

text = raw.decode(
    "utf-8",
    errors="ignore",
)

return text, final_url
```

def html_to_text(source: str) -> str:

```
source = re.sub(
    r"(?is)<script.*?>.*?</script>",
    " ",
    source,
)

source = re.sub(
    r"(?is)<style.*?>.*?</style>",
    " ",
    source,
)

source = re.sub(
    r"(?is)<noscript.*?>.*?</noscript>",
    " ",
    source,
)

source = re.sub(
    r"<[^>]+>",
    " ",
    source,
)

source = html.unescape(source)

return clean_text(source)
```

def extract_title(source: str) -> str:

```
match = re.search(
    r"(?is)<title[^>]*>(.*?)</title>",
    source,
)

if not match:
    return ""

return clean_text(
    html.unescape(match.group(1))
)
```

def extract_prudential_research_metadata(
source: str,
) -> dict:

```
text = html_to_text(source)

result = {
    "underlyingFund": "",
    "underlyingManager": "",
}

patterns = [
    (
        "underlyingFund",
        [
            r"feeds\s+(?:100%\s+)?into\s+(.*?)(?:\s+\(the\s+Underlying Fund\)|,?\s+which aims)",
            r"Underlying Fund\s*[:\-]\s*(.*?)(?:\s+Investment manager|\s+Manager)",
        ],
    ),
    (
        "underlyingManager",
        [
            r"Investment manager of the underlying fund\s+(.*?)(?:\s+Payout frequency|\s+Continuing Investment Charge|\s+Fund documents|$)",
            r"Investment manager of the Underlying Fund\s+(.*?)(?:\s+Fund documents|$)",
        ],
    ),
]

for field, field_patterns in patterns:

    for pattern in field_patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:

            value = clean_text(
                match.group(1)
            )

            value = re.sub(
                r"\s+",
                " ",
                value,
            ).strip(
                " .,:;-"
            )

            if value:
                result[field] = value
                break

return result
```

def extract_links(
source: str,
base_url: str,
) -> list[str]:

```
links = []

for match in re.finditer(
    r"""(?is)<a\b[^>]*href\s*=\s*["']([^"']+)["']""",
    source,
):

    href = html.unescape(
        match.group(1)
    ).strip()

    if not href:
        continue

    absolute = urljoin(
        base_url,
        href,
    )

    parsed = urlparse(
        absolute
    )

    if parsed.scheme not in {
        "http",
        "https",
    }:
        continue

    if parsed.path.lower().endswith(".pdf"):
        continue

    links.append(
        absolute
    )

return list(
    dict.fromkeys(links)
)
```

def domain_matches_manager(
url: str,
manager: str,
) -> bool:

```
hostname = (
    urlparse(url)
    .hostname
    or ""
).lower()

manager_text = clean_text(
    manager
).lower()

if not hostname or not manager_text:
    return False

manager_words = re.findall(
    r"[a-z0-9]+",
    manager_text,
)

important = [
    word
    for word in manager_words
    if len(word) >= 5
    and word not in {
        "investments",
        "investment",
        "management",
        "limited",
        "company",
        "singapore",
        "asset",
        "assets",
        "fund",
        "funds",
    }
]

return any(
    word in hostname
    for word in important
)
```

# =============================================================================

# SEARCH ENGINE DISCOVERY

# =============================================================================

def extract_bing_result_urls(
source: str,
) -> list[str]:

```
urls = []

patterns = [
    r'<a[^>]+href="(https?://[^"]+)"',
    r'<a[^>]+href=\'(https?://[^\']+)\'',
]

for pattern in patterns:

    for match in re.finditer(
        pattern,
        source,
        flags=re.IGNORECASE,
    ):

        url = html.unescape(
            match.group(1)
        )

        if "bing.com" in (
            urlparse(url)
            .hostname
            or ""
        ).lower():
            continue

        if url.lower().split("?")[0].endswith(
            ".pdf"
        ):
            continue

        urls.append(url)

return list(
    dict.fromkeys(urls)
)
```

def search_web(
query: str,
) -> list[str]:

```
encoded = quote_plus(query)

url = (
    "https://www.bing.com/search?"
    f"q={encoded}&count=10"
)

try:

    source, _ = fetch_html(
        url,
        timeout=RESEARCH_TIMEOUT_SECONDS,
    )

    return extract_bing_result_urls(
        source
    )

except Exception as error:

    print(
        f"WARNING: Web search failed for "
        f"{query!r}: {error}"
    )

    return []
```

# =============================================================================

# ONLINE ALLOCATION PARSING

# =============================================================================

def parse_percentage(
value: str,
) -> float | None:

```
if value is None:
    return None

text = clean_text(value)

match = re.search(
    r"(-?\d+(?:\.\d+)?)\s*%",
    text,
)

if not match:
    return None

try:
    number = float(
        match.group(1)
    )
except ValueError:
    return None

if not 0 <= number <= 100:
    return None

return number
```

def normalize_country(
value: str,
) -> str:

```
value = clean_text(value)

aliases = {
    "USA": "United States",
    "US": "United States",
    "United States of America": "United States",
    "UK": "United Kingdom",
    "Korea, Republic of": "South Korea",
    "Republic of Korea": "South Korea",
    "Taiwan (Republic of China)": "Taiwan",
}

return aliases.get(
    value,
    value,
)
```

def region_for_country(
country: str,
) -> str | None:

```
normalized = normalize_country(
    country
)

return COUNTRY_REGIONS.get(
    normalized
)
```

def parse_html_tables(
source: str,
) -> list[list[list[str]]]:

```
tables = []

for table_match in re.finditer(
    r"(?is)<table\b[^>]*>(.*?)</table>",
    source,
):

    table_html = table_match.group(1)

    rows = []

    for row_match in re.finditer(
        r"(?is)<tr\b[^>]*>(.*?)</tr>",
        table_html,
    ):

        row_html = row_match.group(1)

        cells = []

        for cell_match in re.finditer(
            r"(?is)<(?:td|th)\b[^>]*>(.*?)</(?:td|th)>",
            row_html,
        ):

            cell = html_to_text(
                cell_match.group(1)
            )

            cells.append(cell)

        if cells:
            rows.append(cells)

    if rows:
        tables.append(rows)

return tables
```

def table_kind(
rows: list[list[str]],
) -> str | None:

```
sample = " ".join(
    " ".join(row)
    for row in rows[:5]
).lower()

if any(
    term in sample
    for term in [
        "sector allocation",
        "sector",
        "industry allocation",
        "industry",
    ]
):
    return "sector"

if any(
    term in sample
    for term in [
        "country allocation",
        "market allocation",
        "geographic allocation",
        "geographical allocation",
        "country",
        "market",
    ]
):
    return "country"

return None
```

def parse_allocation_tables(
source: str,
) -> tuple[list[dict], list[dict]]:

```
country_rows = []
sector_rows = []

for rows in parse_html_tables(source):

    kind = table_kind(rows)

    if kind is None:
        continue

    for row in rows:

        if len(row) < 2:
            continue

        weight = None
        label = ""

        for cell in row:

            candidate = parse_percentage(
                cell
            )

            if candidate is not None:
                weight = candidate
                break

        if weight is None:
            continue

        for cell in row:

            if parse_percentage(cell) is not None:
                continue

            candidate = clean_text(
                cell
            )

            if candidate:
                label = candidate
                break

        if not label:
            continue

        item = {
            "name": label,
            "weight": weight,
        }

        if kind == "country":
            country_rows.append(item)

        elif kind == "sector":
            sector_rows.append(item)

return (
    country_rows,
    sector_rows,
)
```

def parse_allocation_from_text(
source: str,
) -> tuple[list[dict], list[dict]]:

```
text = html_to_text(source)

country_rows = []
sector_rows = []

country_names = sorted(
    COUNTRY_REGIONS.keys(),
    key=len,
    reverse=True,
)

for country in country_names:

    pattern = (
        rf"\b{re.escape(country)}\b"
        rf".{{0,100}}?"
        rf"(\d+(?:\.\d+)?)\s*%"
    )

    match = re.search(
        pattern,
        text,
        flags=re.IGNORECASE,
    )

    if match:

        try:
            weight = float(
                match.group(1)
            )
        except ValueError:
            continue

        if 0 <= weight <= 100:

            country_rows.append(
                {
                    "name": normalize_country(
                        country
                    ),
                    "weight": weight,
                }
            )

sector_names = [
    "Information Technology",
    "Financials",
    "Industrials",
    "Consumer Discretionary",
    "Consumer Staples",
    "Health Care",
    "Health Care",
    "Communication Services",
    "Energy",
    "Materials",
    "Utilities",
    "Real Estate",
]

for sector in sector_names:

    pattern = (
        rf"\b{re.escape(sector)}\b"
        rf".{{0,100}}?"
        rf"(\d+(?:\.\d+)?)\s*%"
    )

    match = re.search(
        pattern,
        text,
        flags=re.IGNORECASE,
    )

    if match:

        try:
            weight = float(
                match.group(1)
            )
        except ValueError:
            continue

        if 0 <= weight <= 100:

            sector_rows.append(
                {
                    "name": sector,
                    "weight": weight,
                }
            )

return (
    country_rows,
    sector_rows,
)
```

# =============================================================================

# GEOGRAPHIC / SECTOR OUTPUT

# =============================================================================

def build_geographic_exposure(
country_rows: list[dict],
) -> list[dict]:

```
region_countries: dict[str, dict[str, float]] = (
    defaultdict(dict)
)

for item in country_rows:

    country = normalize_country(
        item.get("name", "")
    )

    weight = item.get(
        "weight"
    )

    if not country:
        continue

    if (
        isinstance(weight, bool)
        or not isinstance(
            weight,
            (int, float),
        )
        or not 0 <= weight <= 100
    ):
        continue

    region = region_for_country(
        country
    )

    if region is None:
        continue

    # If duplicate country entries occur,
    # retain the highest published weight.
    previous = region_countries[
        region
    ].get(
        country
    )

    if (
        previous is None
        or weight > previous
    ):
        region_countries[
            region
        ][country] = float(weight)

ranked_regions = []

for region, countries in (
    region_countries.items()
):

    total = sum(
        countries.values()
    )

    ranked_regions.append(
        (
            total,
            region,
            countries,
        )
    )

ranked_regions.sort(
    key=lambda item: (
        -item[0],
        item[1],
    )
)

output = []

for total, region, countries in (
    ranked_regions[
        :MAX_RESEARCH_REGIONS
    ]
):

    ordered_countries = sorted(
        countries.items(),
        key=lambda item: (
            -item[1],
            item[0],
        ),
    )

    output.append(
        {
            "region": region,
            "countries": [
                {
                    "country": country,
                    "weight": weight,
                }
                for country, weight in (
                    ordered_countries
                )
            ],
        }
    )

return output
```

def build_top_sectors(
sector_rows: list[dict],
) -> list[dict]:

```
sectors = {}

for item in sector_rows:

    sector = clean_text(
        item.get("name")
    )

    weight = item.get(
        "weight"
    )

    if not sector:
        continue

    if (
        isinstance(weight, bool)
        or not isinstance(
            weight,
            (int, float),
        )
        or not 0 <= weight <= 100
    ):
        continue

    previous = sectors.get(
        sector
    )

    if (
        previous is None
        or weight > previous
    ):
        sectors[sector] = float(
            weight
        )

ordered = sorted(
    sectors.items(),
    key=lambda item: (
        -item[1],
        item[0],
    ),
)

return [
    {
        "sector": sector,
        "weight": weight,
    }
    for sector, weight in ordered[
        :MAX_RESEARCH_SECTORS
    ]
]
```

# =============================================================================

# ONLINE FUND RESEARCH

# =============================================================================

def research_fund_online(
prudential_url: str,
prudential: dict,
) -> dict:

```
result = {
    "status": "unresolved",
    "sourceType": "none",
    "underlyingFund": "",
    "underlyingManager": "",
    "sourceUrl": None,
    "sourceDate": None,
    "geographicExposure": [],
    "topSectors": [],
}

# -------------------------------------------------------------------------
# PRUDENTIAL ONLINE PAGE
# -------------------------------------------------------------------------

try:

    prudential_html, final_prudential_url = (
        fetch_html(
            prudential_url
        )
    )

except Exception as error:

    result["error"] = (
        f"Prudential online page failed: {error}"
    )

    return result

metadata = (
    extract_prudential_research_metadata(
        prudential_html
    )
)

underlying_fund = clean_text(
    metadata.get(
        "underlyingFund"
    )
)

underlying_manager = clean_text(
    metadata.get(
        "underlyingManager"
    )
)

result["underlyingFund"] = (
    underlying_fund
)

result["underlyingManager"] = (
    underlying_manager
)

if not underlying_fund:

    result["status"] = "partial"
    result["sourceType"] = "prudential_online"
    result["sourceUrl"] = (
        final_prudential_url
    )
    result["sourceDate"] = (
        singapore_today()
    )
    result["error"] = (
        "Underlying fund could not be identified "
        "from Prudential online page."
    )

    return result

# -------------------------------------------------------------------------
# SEARCH FOR MANAGER ONLINE PAGE
# -------------------------------------------------------------------------

manager_candidates = []

# First inspect links already present on
# the Prudential page.
for link in extract_links(
    prudential_html,
    final_prudential_url,
):

    if domain_matches_manager(
        link,
        underlying_manager,
    ):

        manager_candidates.append(
            link
        )

# Then use online search.
search_queries = [
    f'"{underlying_fund}"',
    (
        f'"{underlying_fund}" '
        f'"{underlying_manager}"'
    ),
    (
        f'"{underlying_fund}" '
        f'country sector allocation'
    ),
]

for query in search_queries:

    if len(manager_candidates) >= 10:
        break

    for link in search_web(
        query
    ):

        if domain_matches_manager(
            link,
            underlying_manager,
        ):
            manager_candidates.append(
                link
            )

manager_candidates = list(
    dict.fromkeys(
        manager_candidates
    )
)

if not manager_candidates:

    # Secondary search without manager
    # domain restriction.
    for query in search_queries:

        for link in search_web(
            query
        ):

            if link.lower().split(
                "?"
            )[0].endswith(".pdf"):
                continue

            manager_candidates.append(
                link
            )

        if manager_candidates:
            break

manager_candidates = list(
    dict.fromkeys(
        manager_candidates
    )
)

# -------------------------------------------------------------------------
# FETCH CANDIDATE HTML SOURCES
# -------------------------------------------------------------------------

for candidate in manager_candidates[
    :12
]:

    if candidate.lower().split(
        "?"
    )[0].endswith(".pdf"):
        continue

    try:

        time.sleep(
            RESEARCH_DELAY_SECONDS
        )

        candidate_html, final_url = (
            fetch_html(
                candidate
            )
        )

    except Exception:
        continue

    title = extract_title(
        candidate_html
    )

    page_text = html_to_text(
        candidate_html
    )

    fund_match = (
        underlying_fund.lower()
        in (
            title + " " + page_text
        ).lower()
    )

    if not fund_match:

        # Search pages may have redirected
        # to a general manager page.
        # Keep only pages with meaningful
        # allocation terminology.
        allocation_match = any(
            term in page_text.lower()
            for term in [
                "sector allocation",
                "market allocation",
                "country allocation",
                "geographic allocation",
                "geographical allocation",
            ]
        )

        if not allocation_match:
            continue

    country_rows, sector_rows = (
        parse_allocation_tables(
            candidate_html
        )
    )

    if not country_rows and not sector_rows:

        country_rows, sector_rows = (
            parse_allocation_from_text(
                candidate_html
            )
        )

    geographic = (
        build_geographic_exposure(
            country_rows
        )
    )

    sectors = build_top_sectors(
        sector_rows
    )

    if not geographic and not sectors:
        continue

    result["status"] = "success"

    if domain_matches_manager(
        final_url,
        underlying_manager,
    ):
        result["sourceType"] = (
            "official_online"
        )
    else:
        result["sourceType"] = (
            "secondary_online"
        )

    result["sourceUrl"] = (
        final_url
    )

    result["sourceDate"] = (
        singapore_today()
    )

    result["geographicExposure"] = (
        geographic
    )

    result["topSectors"] = sectors

    if not geographic or not sectors:
        result["status"] = "partial"

    return result

result["status"] = "partial"
result["sourceType"] = (
    "prudential_online"
)
result["sourceUrl"] = (
    final_prudential_url
)
result["sourceDate"] = (
    singapore_today()
)
result["error"] = (
    "No usable online HTML allocation source found."
)

return result
```

# =============================================================================

# HOLDINGS

# =============================================================================

def clean_holdings(
holdings,
label: str,
) -> list[dict]:

```
if (
    not isinstance(holdings, list)
    or not holdings
):
    raise ValueError(
        f"{label}: holdings list is empty or invalid."
    )

if len(holdings) > MAX_HOLDINGS:

    raise ValueError(
        f"{label}: more than {MAX_HOLDINGS} holdings."
    )

cleaned = []

for position, item in enumerate(
    holdings,
    start=1,
):

    if not isinstance(item, dict):

        raise ValueError(
            f"{label}: holding {position} "
            f"is not an object."
        )

    name = clean_text(
        item.get("name")
    )

    weight = item.get(
        "weightPercent"
    )

    if item.get("rank") != position:

        raise ValueError(
            f"{label}: rank "
            f"{item.get('rank')} "
            f"at position {position}."
        )

    if not name:

        raise ValueError(
            f"{label}: holding "
            f"{position} has no name."
        )

    if (
        isinstance(weight, bool)
        or not isinstance(
            weight,
            (int, float),
        )
        or not 0 <= weight <= 100
    ):

        raise ValueError(
            f"{label}: holding "
            f"{position} has invalid "
            f"weight {weight!r}."
        )

    cleaned.append(
        {
            "rank": position,
            "name": name,
            "weightPercent": weight,
            "weightText": clean_text(
                item.get(
                    "weightText"
                )
            ),
        }
    )

return cleaned
```

def holdings_block(
result: dict,
stage: str,
label: str,
):

```
status = result.get("status")

common = {
    "source": stage,
    "parser": result.get(
        "holdingsParser"
    ),
    "factsheetUrl": result.get(
        "factsheetUrl"
    ),
    "factsheetDocumentDate": result.get(
        "factsheetDocumentDate"
    ),
    "factsheetDataAsAt": result.get(
        "factsheetDataAsAt"
    ),
}

if status == "success":

    holdings = clean_holdings(
        result.get("topHoldings"),
        label,
    )

    return {
        "status": "published",
        **common,
        "count": len(holdings),
        "holdings": holdings,
    }

if status == "no_holdings_section":

    return {
        "status": "no_holdings_section",
        **common,
        "count": 0,
        "holdings": [],
    }

return None
```

def resolve_holdings(
excel_funds: dict[int, dict],
) -> dict[int, dict]:

```
resolved: dict[int, dict] = {}

def accept(
    result: dict,
    stage: str,
) -> None:

    if result.get("excelRow") is None:

        raise RuntimeError(
            f"{stage}: result without excelRow."
        )

    row = int(
        result["excelRow"]
    )

    if row not in excel_funds:

        raise RuntimeError(
            f"{stage}: row {row} "
            f"is not in Excel."
        )

    if row in resolved:
        return

    result_url = clean_text(
        result.get(
            "prudentialUrl"
        )
    )

    if (
        result_url
        and result_url
        != excel_funds[row][
            "prudentialUrl"
        ]
    ):

        raise RuntimeError(
            f"{stage}: URL mismatch "
            f"for row {row}."
        )

    block = holdings_block(
        result,
        stage,
        f"{stage} row {row}",
    )

    if block is not None:

        block["fundName"] = (
            clean_text(
                result.get(
                    "fundName"
                )
            )
            or None
        )

        resolved[row] = block

if not BASELINE_FILE.exists():

    raise FileNotFoundError(
        f"Baseline output not found: "
        f"{BASELINE_FILE}"
    )

baseline = load_json(
    BASELINE_FILE
)

for result in baseline.get(
    "funds",
    [],
):

    accept(
        result,
        "baseline",
    )

for stage, stage_dir in RECOVERY_STAGES:

    funds_dir = (
        stage_dir / "funds"
    )

    if not funds_dir.exists():

        print(
            f"WARNING: {funds_dir} "
            f"not found; {stage} skipped."
        )

        continue

    for path in sorted(
        funds_dir.glob(
            "*/top_holdings.json"
        )
    ):

        accept(
            load_json(path),
            stage,
        )

return resolved
```

# =============================================================================

# PRUACCESS

# =============================================================================

def validate_bid_history(
history: dict,
label: str,
) -> list[dict]:

```
observations = history.get(
    "observations"
)

if (
    not isinstance(
        observations,
        list,
    )
    or not observations
):

    raise ValueError(
        f"{label}: no BID observations."
    )

if (
    history.get(
        "observationCount"
    )
    != len(observations)
):

    raise ValueError(
        f"{label}: observationCount mismatch."
    )

previous = None
seen = set()

for item in observations:

    date = item.get(
        "date"
    )

    price = item.get(
        "bidPrice"
    )

    if (
        not isinstance(
            date,
            str,
        )
        or not date
    ):

        raise ValueError(
            f"{label}: invalid date "
            f"{date!r}."
        )

    if (
        isinstance(price, bool)
        or not isinstance(
            price,
            (int, float),
        )
    ):

        raise ValueError(
            f"{label}: invalid BID "
            f"price on {date}."
        )

    if date in seen:

        raise ValueError(
            f"{label}: duplicate date "
            f"{date}."
        )

    if (
        previous is not None
        and date < previous
    ):

        raise ValueError(
            f"{label}: observations "
            f"not chronological."
        )

    seen.add(date)
    previous = date

return [
    {
        "date": item["date"],
        "bidPrice": item["bidPrice"],
    }
    for item in observations
]
```

def load_pruaccess() -> dict[int, dict]:

```
result = {}

if not PRUACCESS_FUNDS_DIR.exists():

    print(
        f"WARNING: "
        f"{PRUACCESS_FUNDS_DIR} "
        f"not found."
    )

    return result

for directory in sorted(
    PRUACCESS_FUNDS_DIR.iterdir()
):

    match = re.match(
        r"^(\d+)_",
        directory.name,
    )

    if (
        not directory.is_dir()
        or not match
    ):
        continue

    bid_file = (
        directory
        / "bid_history.json"
    )

    fund_file = (
        directory
        / "prudential_fund.json"
    )

    if not (
        bid_file.exists()
        and fund_file.exists()
    ):
        continue

    row = int(
        match.group(1)
    )

    result[row] = {
        "prudential": load_json(
            fund_file
        ),
        "bidHistory": load_json(
            bid_file
        ),
    }

return result
```

# =============================================================================

# PREVIOUS DATA HELPERS

# =============================================================================

def previous_fund_record(
previous_funds: dict[int, dict],
row: int,
) -> dict | None:

```
record = previous_funds.get(
    row
)

if not isinstance(
    record,
    dict,
):
    return None

return record
```

def previous_bid_record(
previous_bids: dict[int, dict],
row: int,
) -> dict | None:

```
record = previous_bids.get(
    row
)

if not isinstance(
    record,
    dict,
):
    return None

return record
```

def make_retained_fund_record(
previous: dict,
) -> dict:

```
retained = json.loads(
    json.dumps(
        previous,
        ensure_ascii=False,
    )
)

retained["dataStatus"] = (
    "retained_previous"
)

return retained
```

def make_retained_bid_record(
previous: dict,
) -> dict:

```
retained = json.loads(
    json.dumps(
        previous,
        ensure_ascii=False,
    )
)

retained["dataStatus"] = (
    "retained_previous"
)

return retained
```

# =============================================================================

# BUILD

# =============================================================================

def main() -> int:

```
print("=" * 72)
print("VGRAT FMS - BUILD FUNDS DATA")
print("=" * 72)

generated_at = utc_now_iso()
retrieval_date = singapore_today()

print(
    f"Build UTC time:       {generated_at}"
)

print(
    f"Singapore date:       {retrieval_date}"
)

print(
    f"Allow unresolved:     {ALLOW_UNRESOLVED}"
)

# -------------------------------------------------------------------------
# MASTER INPUTS
# -------------------------------------------------------------------------

excel_funds = read_excel_funds()

holdings = resolve_holdings(
    excel_funds
)

pruaccess = load_pruaccess()

# -------------------------------------------------------------------------
# PREVIOUS PUBLISHED DATA
# -------------------------------------------------------------------------

previous_funds_file = load_previous_file(
    FUNDS_OUT
)

previous_bid_file = load_previous_file(
    BID_OUT
)

previous_funds = index_previous_records(
    previous_funds_file
)

previous_bids = index_previous_records(
    previous_bid_file
)

print(
    f"Previous fund records: "
    f"{len(previous_funds)}"
)

print(
    f"Previous BID records:  "
    f"{len(previous_bids)}"
)

# -------------------------------------------------------------------------
# OUTPUT COLLECTIONS
# -------------------------------------------------------------------------

fund_records = []
bid_records = []

# -------------------------------------------------------------------------
# COUNTERS
# -------------------------------------------------------------------------

fund_updated = 0
fund_inserted = 0
fund_retained = 0
fund_unavailable = 0

bid_updated = 0
bid_inserted = 0
bid_retained = 0
bid_unavailable = 0

holdings_published = 0
holdings_no_section = 0
holdings_retained = 0
holdings_unresolved = 0

stage_counts: dict[str, int] = {}

total_observations = 0

dividend_true = 0
dividend_false = 0

research_success = 0
research_partial = 0
research_unresolved = 0
research_geography = 0
research_sectors = 0

# -------------------------------------------------------------------------
# PROCESS EVERY EXCEL FUND INDEPENDENTLY
# -------------------------------------------------------------------------

for row in sorted(excel_funds):

    excel = excel_funds[row]

    pru = pruaccess.get(
        row
    )

    prudential = (
        (pru or {}).get(
            "prudential"
        )
        or {}
    )

    previous_fund = (
        previous_fund_record(
            previous_funds,
            row,
        )
    )

    previous_bid = (
        previous_bid_record(
            previous_bids,
            row,
        )
    )

    # =====================================================================
    # IDENTITY
    # =====================================================================

    identity = {
        "excelRow": row,

        "fundIdentifier": (
            clean_text(
                prudential.get(
                    "fundIdentifier"
                )
            )
            or None
        ),

        "fundCode": (
            clean_text(
                prudential.get(
                    "fundCode"
                )
            )
            or None
        ),

        "fundName": (
            clean_text(
                prudential.get(
                    "fundName"
                )
            )
            or None
        ),

        "pruAccessName": (
            excel[
                "pruAccessName"
            ]
            or None
        ),

        "prudentialUrl": (
            excel[
                "prudentialUrl"
            ]
        ),
    }

    # =====================================================================
    # CURRENT HOLDINGS
    # =====================================================================

    block = holdings.get(
        row
    )

    current_holdings_success = (
        block is not None
    )

    # =====================================================================
    # CURRENT FUND INFORMATION
    # =====================================================================

    current_fund_success = bool(
        prudential
    )

    # =====================================================================
    # CURRENT BID HISTORY
    # =====================================================================

    current_bid_success = False
    current_observations = None

    if pru is not None:

        try:

            current_observations = (
                validate_bid_history(
                    pru[
                        "bidHistory"
                    ],
                    f"row {row} BID",
                )
            )

            current_bid_success = True

        except Exception as error:

            print(
                f"WARNING: Row {row} "
                f"BID failed: {error}"
            )

    # =====================================================================
    # CORE FUND SUCCESS
    # =====================================================================

    complete_success = (
        current_fund_success
        and current_holdings_success
        and current_bid_success
    )

    # =====================================================================
    # SUCCESSFUL FUND
    # =====================================================================

    if complete_success:

        # -----------------------------------------------------------------
        # FUND INFO
        # -----------------------------------------------------------------

        fund_info = {
            k: v
            for k, v in prudential.items()
            if k not in {
                "raw",
                "fundIdentifier",
                "fundCode",
                "fundName",
            }
        }

        fund_info = (
            normalize_dividend_fields(
                fund_info
            )
        )

        if fund_info[
            "hasDividend"
        ]:
            dividend_true += 1
        else:
            dividend_false += 1

        # -----------------------------------------------------------------
        # ONLINE RESEARCH
        # -----------------------------------------------------------------
        #
        # Research is deliberately independent from complete_success.
        # A research failure does NOT make the fund fail.
        # -----------------------------------------------------------------

        try:

            research = (
                research_fund_online(
                    prudential_url=(
                        excel[
                            "prudentialUrl"
                        ]
                    ),
                    prudential=prudential,
                )
            )

        except Exception as error:

            research = {
                "status": "unresolved",
                "sourceType": "none",
                "underlyingFund": "",
                "underlyingManager": "",
                "sourceUrl": None,
                "sourceDate": None,
                "geographicExposure": [],
                "topSectors": [],
                "error": str(error),
            }

        research_status = research.get(
            "status"
        )

        if research_status == "success":
            research_success += 1

        elif research_status == "partial":
            research_partial += 1

        else:
            research_unresolved += 1

        if research.get(
            "geographicExposure"
        ):
            research_geography += 1

        if research.get(
            "topSectors"
        ):
            research_sectors += 1

        # -----------------------------------------------------------------
        # HOLDINGS
        # -----------------------------------------------------------------

        top_holdings = {
            k: v
            for k, v in block.items()
            if k != "fundName"
        }

        stage = block[
            "source"
        ]

        stage_counts[stage] = (
            stage_counts.get(
                stage,
                0,
            )
            \+ 1
        )

        if (
            block[
                "status"
            ]
            == "published"
        ):
            holdings_published += 1

        else:
            holdings_no_section += 1

        # -----------------------------------------------------------------
        # DATA STATUS
        # -----------------------------------------------------------------

        if previous_fund is None:

            data_status = "inserted"
            fund_inserted += 1

        else:

            data_status = "updated"
            fund_updated += 1

        # -----------------------------------------------------------------
        # FINAL FUND RECORD
        # -----------------------------------------------------------------

        fund_record = {
            **identity,

            "dataStatus": data_status,

            "dataRetrievedDate": (
                retrieval_date
            ),

            "fund": fund_info,

            "geographicExposure": (
                research.get(
                    "geographicExposure",
                    [],
                )
            ),

            "topSectors": (
                research.get(
                    "topSectors",
                    [],
                )
            ),

            "research": {
                k: v
                for k, v in research.items()
                if k not in {
                    "geographicExposure",
                    "topSectors",
                }
            },

            "topHoldings": top_holdings,
        }

        fund_records.append(
            fund_record
        )

        # -----------------------------------------------------------------
        # BID
        # -----------------------------------------------------------------

        history = pru[
            "bidHistory"
        ]

        current_bid_history = {
            "status": "success",

            "priceType": "BID",

            "pruAccessFundId": (
                history.get(
                    "fundId"
                )
            ),

            "currency": (
                history.get(
                    "currency"
                )
            ),

            "startDate": (
                history.get(
                    "startDate"
                )
            ),

            "endDate": (
                history.get(
                    "endDate"
                )
            ),

            "observationCount": len(
                current_observations
            ),

            "observations": (
                current_observations
            ),
        }

        if previous_bid is None:

            bid_status = "inserted"
            bid_inserted += 1

        else:

            bid_status = "updated"
            bid_updated += 1

        bid_records.append(
            {
                **identity,

                "dataStatus": bid_status,

                "dataRetrievedDate": (
                    retrieval_date
                ),

                "bidHistory": (
                    current_bid_history
                ),
            }
        )

        total_observations += len(
            current_observations
        )

        print(
            f"Row {row}: SUCCESS "
            f"({data_status}) "
            f"research={research_status}"
        )

        continue

    # =====================================================================
    # FAILED CURRENT FUND
    # =====================================================================

    failure_reasons = []

    if not current_fund_success:
        failure_reasons.append(
            "fund data"
        )

    if not current_holdings_success:
        failure_reasons.append(
            "holdings"
        )

    if not current_bid_success:
        failure_reasons.append(
            "BID history"
        )

    failure_text = ", ".join(
        failure_reasons
    )

    # =====================================================================
    # PREVIOUS FUND EXISTS -> RETAIN
    # =====================================================================

    if previous_fund is not None:

        retained_fund = (
            make_retained_fund_record(
                previous_fund
            )
        )

        fund_records.append(
            retained_fund
        )

        fund_retained += 1

        print(
            f"Row {row}: FAILED "
            f"({failure_text}) -> "
            f"RETAINED PREVIOUS FUND"
        )

    else:

        fund_unavailable += 1

        print(
            f"Row {row}: FAILED "
            f"({failure_text}) -> "
            f"NO PREVIOUS FUND; "
            f"NOT PUBLISHED"
        )

    # =====================================================================
    # BID RETENTION IS INDEPENDENT
    # =====================================================================

    if current_bid_success:

        history = pru[
            "bidHistory"
        ]

        bid_history = {
            "status": "success",

            "priceType": "BID",

            "pruAccessFundId": (
                history.get(
                    "fundId"
                )
            ),

            "currency": (
                history.get(
                    "currency"
                )
            ),

            "startDate": (
                history.get(
                    "startDate"
                )
            ),

            "endDate": (
                history.get(
                    "endDate"
                )
            ),

            "observationCount": len(
                current_observations
            ),

            "observations": (
                current_observations
            ),
        }

        if previous_bid is None:

            bid_status = "inserted"
            bid_inserted += 1

        else:

            bid_status = "updated"
            bid_updated += 1

        bid_records.append(
            {
                **identity,

                "dataStatus": bid_status,

                "dataRetrievedDate": (
                    retrieval_date
                ),

                "bidHistory": bid_history,
            }
        )

        total_observations += len(
            current_observations
        )

    elif previous_bid is not None:

        retained_bid = (
            make_retained_bid_record(
                previous_bid
            )
        )

        bid_records.append(
            retained_bid
        )

        bid_retained += 1

        print(
            f"Row {row}: BID FAILED -> "
            f"RETAINED PREVIOUS BID"
        )

    else:

        bid_unavailable += 1

        print(
            f"Row {row}: BID FAILED -> "
            f"NO PREVIOUS BID; "
            f"NOT PUBLISHED"
        )

# =============================================================================
# BUILD ENVELOPES
# =============================================================================

fund_envelope = {
    "schemaVersion": SCHEMA_VERSION,

    "generatedAtUtc": generated_at,

    "dataRetrievedDate": retrieval_date,

    "source": str(
        EXCEL_FILE
    ),

    "fundCount": len(
        fund_records
    ),

    "summary": {
        "fundsUpdated": fund_updated,

        "fundsInserted": fund_inserted,

        "fundsRetainedPrevious": (
            fund_retained
        ),

        "fundsUnavailable": (
            fund_unavailable
        ),

        "holdingsPublished": (
            holdings_published
        ),

        "noHoldingsSection": (
            holdings_no_section
        ),

        "holdingsRetainedPrevious": (
            holdings_retained
        ),

        "holdingsUnresolved": (
            holdings_unresolved
        ),

        "holdingsBySource": (
            stage_counts
        ),

        "fundsWithDividend": (
            dividend_true
        ),

        "fundsWithoutDividend": (
            dividend_false
        ),

        "researchSuccess": (
            research_success
        ),

        "researchPartial": (
            research_partial
        ),

        "researchUnresolved": (
            research_unresolved
        ),

        "researchWithGeographicExposure": (
            research_geography
        ),

        "researchWithTopSectors": (
            research_sectors
        ),
    },

    "funds": fund_records,
}

bid_envelope = {
    "schemaVersion": SCHEMA_VERSION,

    "generatedAtUtc": generated_at,

    "dataRetrievedDate": retrieval_date,

    "source": str(
        EXCEL_FILE
    ),

    "fundCount": len(
        bid_records
    ),

    "summary": {
        "fundsUpdated": bid_updated,

        "fundsInserted": bid_inserted,

        "fundsRetainedPrevious": (
            bid_retained
        ),

        "fundsUnavailable": (
            bid_unavailable
        ),

        "totalObservations": (
            total_observations
        ),
    },

    "funds": bid_records,
}

# =============================================================================
# WRITE
# =============================================================================

write_json_atomic(
    FUNDS_OUT,
    fund_envelope,
)

write_json_atomic(
    BID_OUT,
    bid_envelope,
    compact=True,
)

# =============================================================================
# SUMMARY
# =============================================================================

print("\n" + "=" * 72)
print("BUILD COMPLETE")
print("=" * 72)

print(
    f"Excel universe:          "
    f"{len(excel_funds)}"
)

print(
    f"Published fund records:  "
    f"{len(fund_records)}"
)

print(
    f"Funds updated:           "
    f"{fund_updated}"
)

print(
    f"Funds inserted:          "
    f"{fund_inserted}"
)

print(
    f"Funds retained:          "
    f"{fund_retained}"
)

print(
    f"Funds unavailable:       "
    f"{fund_unavailable}"
)

print(
    f"Holdings published:      "
    f"{holdings_published}"
)

print(
    f"No holdings section:     "
    f"{holdings_no_section}"
)

print(
    f"Holdings by source:      "
    f"{stage_counts}"
)

print(
    f"BID updated:             "
    f"{bid_updated}"
)

print(
    f"BID inserted:            "
    f"{bid_inserted}"
)

print(
    f"BID retained:            "
    f"{bid_retained}"
)

print(
    f"BID unavailable:         "
    f"{bid_unavailable}"
)

print(
    f"BID observations:        "
    f"{total_observations}"
)

print(
    f"Funds with dividend:     "
    f"{dividend_true}"
)

print(
    f"Funds without dividend:  "
    f"{dividend_false}"
)

print(
    f"Research success:        "
    f"{research_success}"
)

print(
    f"Research partial:        "
    f"{research_partial}"
)

print(
    f"Research unresolved:     "
    f"{research_unresolved}"
)

print(
    f"Research geography:      "
    f"{research_geography}"
)

print(
    f"Research sectors:        "
    f"{research_sectors}"
)

print(
    f"Singapore retrieval date:"
    f" {retrieval_date}"
)

print(
    f"Generated UTC:            "
    f"{generated_at}"
)

print(
    f"Wrote: {FUNDS_OUT}"
)

print(
    f"Wrote: {BID_OUT}"
)

return 0
```

if **name** == "**main**":

```
try:

    raise SystemExit(
        main()
    )

except Exception as error:

    print(
        f"\nFATAL BUILD ERROR: {error}",
        file=sys.stderr,
    )

    raise SystemExit(1)
