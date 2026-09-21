```python
#!/usr/bin/env python3

"""
Prudential Singapore official fund Top Holdings extractor.

MASTER SOURCE
=============

Excel:
    Funds Links.xlsm

Column A:
    Prudential Singapore fund URL.

Column B:
    Exact PruAccess fund name.

SCOPE
=====

This script processes every populated URL in Excel Column A.

For each fund it:

1. Opens the official Prudential Singapore fund page.
2. Finds the official Prudential factsheet link.
3. Downloads the official Prudential factsheet PDF.
4. Extracts the "Top 10 Holdings" section.
5. Parses up to 10 holdings.
6. Preserves the exact published holding order.
7. Preserves multiline holding names.
8. Preserves published percentages when available.
9. Stores null when Prudential publishes a holding without a percentage.
10. Performs strict validation.
11. Saves the raw PDF, extracted text, JSON and metadata.
12. Produces a run summary.

HARD RULES
==========

- Excel Column A controls the fund universe.
- No hardcoded fund count.
- Only official Prudential Singapore sources.
- No third-party holdings sources.
- No inferred holding names.
- No fabricated holding names.
- No calculated holding percentages.
- No inferred holding percentages.
- No forced 10 holdings.
- If Prudential publishes fewer than 10 holdings, store exactly that count.
- Holdings remain in Prudential's published order.
- Multiline holding names are preserved/joined.
- Published percentage is optional.
- Missing published percentage = null.
- Rank is the primary holding-row boundary.
- Percentage does NOT define the holding boundary.
- Duplicate percentages are allowed.
- Duplicate holding names are rejected.
- If a Top Holdings section exists but holdings cannot be identified reliably,
  the fund fails.
- If no Top Holdings section exists, the fund is classified as
  "no_holdings_section".
- No synthetic data.
- No interpolation.
- No estimates.
- No carry-forward.
- This script does not modify PruAccess extraction.
- This script does not create data.json.
- This script does not modify frontend files.
- This is a holdings collector/validator only.

IMPORTANT PARSING RULE
======================

The previous parser assumed:

    percentage -> holding boundary

That is unsafe.

Official Prudential factsheets may contain:

    1
    Holding Name

    2
    Another Holding

with no percentages published.

Therefore:

    rank -> holding boundary

Percentage is optional metadata belonging to the holding currently
being parsed.

If a percentage is not explicitly published by Prudential:

    weightPercent = null
    weightText = null

No percentage is ever calculated.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urljoin, urlparse

import openpyxl
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")

MAX_HOLDINGS = 10

REQUEST_TIMEOUT_SECONDS = 60

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
```
