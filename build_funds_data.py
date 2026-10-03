#!/usr/bin/env python3

"""
VGrat FMS - BUILD FUNDS DATA

Merges the outputs of the extraction pipelines into two data files that
share the SAME layout (same envelope, same per-fund identity block):

    data/funds.json
    data/bid_history.json


MASTER UNIVERSE
===============

Funds Links.xlsm, Column A (Prudential URL) / Column B (PruAccess name).

Every populated row in Funds Links.xlsm becomes exactly one record in
BOTH output files, ordered by Excel row.

Funds Links.xlsm is the MASTER UNIVERSE.

Research Funds.xlsx is an EXISTING RESEARCH INPUT only.

No AI research, web research, Morningstar research, factsheet research,
or additional classification is performed by this script.


RESEARCH INPUT
==============

Research workbook:

    Research Funds.xlsx

Worksheet:

    Fund Research

Matching rule:

    Funds Links.xlsm Column B
        exact match
    ->
    Research Funds.xlsx Fund Research
        fund-name column

Research fields extracted:

    Geographic 1
    Geographic 2
    Sector 1
    Sector 2

Research data is copied exactly from the research workbook after
whitespace normalization.

The research workbook does NOT determine the master fund universe.

A research workbook row cannot create a new fund.

A missing research match does NOT make the core fund unresolved.


HOLDINGS RESOLUTION
===================

Stage order:

    baseline -> recovery1 -> recovery2 -> recovery3

- A fund that succeeded in the baseline keeps its baseline result.
- A fund marked "no_holdings_section" remains as-is.
- A fund that failed in baseline is replaced by the first recovery
  stage that resolves it.
- A fund still unresolved after Recovery 3 is "unresolved".

Nothing is inferred or fabricated.


BID HISTORY
===========

Read from:

    output_pruaccess/funds/<row>_<id>/

Files:

    bid_history.json
    prudential_fund.json

Observations are copied exactly as extracted.


DIVIDEND NORMALIZATION
======================

If dividendRate exists and is non-empty:

    hasDividend = true

Existing dividendRate and dividendUnits values are not modified.


OUTPUT
======

data/funds.json
data/bid_history.json
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook


# =============================================================================
# CONFIGURATION
# =============================================================================

MASTER_EXCEL_FILE = Path("Funds Links.xlsm")

RESEARCH_EXCEL_FILE = Path("Research Funds.xlsx")
RESEARCH_SHEET_NAME = "Fund Research"

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

ALLOW_UNRESOLVED = os.environ.get(
    "ALLOW_UNRESOLVED", ""
).strip().lower() in {"1", "true", "yes"}


# =============================================================================
# HELPERS
# =============================================================================

def clean_text(value) -> str:
    """
    Normalize whitespace without otherwise changing the value.
    """
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\xa0", " ")
    ).strip()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(
        microsecond=0
    ).isoformat()


def load_json(path: Path):
    try:
        return json.loads(
            path.read_text(encoding="utf-8")
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
    path.parent.mkdir(
        parents=True,
        exist_ok=True
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
        encoding="utf-8"
    )

    tmp.replace(path)


def normalize_match_key(value) -> str:
    """
    Matching normalization.

    The actual fund name remains unchanged in the output.

    Matching is case-insensitive and whitespace-normalized so that
    accidental Excel spacing/case differences do not prevent a match.

    No fuzzy matching is performed.
    """
    return clean_text(value).casefold()


# =============================================================================
# MASTER EXCEL UNIVERSE
# =============================================================================

def read_excel_funds() -> dict[int, dict]:
    """
    Read Funds Links.xlsm.

    Column A:
        Prudential URL

    Column B:
        PruAccess Fund Name

    Every populated Column A row is part of the master universe.
    """

    if not MASTER_EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Master Excel file not found: {MASTER_EXCEL_FILE}"
        )

    workbook = load_workbook(
        MASTER_EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    funds = {}

    try:
        worksheet = workbook.active

        for row in range(
            2,
            worksheet.max_row + 1
        ):
            url = clean_text(
                worksheet.cell(
                    row=row,
                    column=1
                ).value
            )

            name = clean_text(
                worksheet.cell(
                    row=row,
                    column=2
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
            "No populated URLs found in "
            "Funds Links.xlsm Column A."
        )

    return funds


# =============================================================================
# RESEARCH FUNDS.XLSX
# =============================================================================

def find_research_headers(
    worksheet
) -> dict[str, int]:
    """
    Find the required research headers from the first row.

    Required:

        Fund Name
        Geographic 1
        Geographic 2
        Sector 1
        Sector 2

    The fund-name header accepts the explicit research workbook
    naming variants used by the project, while the four research
    fields remain exact.
    """

    headers = {}

    for column in range(
        1,
        worksheet.max_column + 1
    ):
        value = clean_text(
            worksheet.cell(
                row=1,
                column=column
            ).value
        )

        if value:
            headers[
                normalize_match_key(value)
            ] = column

    fund_name_candidates = [
        "fund name",
        "pruaccess fund name",
        "fund",
        "name",
    ]

    fund_name_column = None

    for candidate in fund_name_candidates:
        column = headers.get(
            normalize_match_key(candidate)
        )

        if column is not None:
            fund_name_column = column
            break

    if fund_name_column is None:
        raise RuntimeError(
            "Research Funds.xlsx / "
            f"'{RESEARCH_SHEET_NAME}' is missing "
            "the fund-name column."
        )

    required_columns = {
        "Geographic 1": None,
        "Geographic 2": None,
        "Sector 1": None,
        "Sector 2": None,
    }

    for header in required_columns:
        column = headers.get(
            normalize_match_key(header)
        )

        if column is None:
            raise RuntimeError(
                "Research Funds.xlsx / "
                f"'{RESEARCH_SHEET_NAME}' is missing "
                f"required column: {header}"
            )

        required_columns[header] = column

    return {
        "fundName": fund_name_column,
        **required_columns,
    }


def read_research_funds() -> dict[str, dict]:
    """
    Read the already-researched Research Funds.xlsx workbook.

    Returns:

        normalized PruAccess fund name ->
        research fields

    No research is performed here.
    """

    if not RESEARCH_EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Research workbook not found: "
            f"{RESEARCH_EXCEL_FILE}"
        )

    workbook = load_workbook(
        RESEARCH_EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    research = {}

    try:
        if RESEARCH_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Research workbook does not contain "
                f"worksheet '{RESEARCH_SHEET_NAME}'. "
                f"Available sheets: "
                f"{', '.join(workbook.sheetnames)}"
            )

        worksheet = workbook[
            RESEARCH_SHEET_NAME
        ]

        columns = find_research_headers(
            worksheet
        )

        for row in range(
            2,
            worksheet.max_row + 1
        ):
            fund_name = clean_text(
                worksheet.cell(
                    row=row,
                    column=columns["fundName"]
                ).value
            )

            if not fund_name:
                continue

            key = normalize_match_key(
                fund_name
            )

            record = {
                "researchFundName": fund_name,
                "geographic1": clean_text(
                    worksheet.cell(
                        row=row,
                        column=columns["Geographic 1"]
                    ).value
                ),
                "geographic2": clean_text(
                    worksheet.cell(
                        row=row,
                        column=columns["Geographic 2"]
                    ).value
                ),
                "sector1": clean_text(
                    worksheet.cell(
                        row=row,
                        column=columns["Sector 1"]
                    ).value
                ),
                "sector2": clean_text(
                    worksheet.cell(
                        row=row,
                        column=columns["Sector 2"]
                    ).value
                ),
            }

            if key in research:
                raise RuntimeError(
                    "Duplicate research fund name found "
                    f"in Research Funds.xlsx: "
                    f"{fund_name!r}"
                )

            research[key] = record

    finally:
        workbook.close()

    if not research:
        raise RuntimeError(
            "No research records found in "
            f"{RESEARCH_EXCEL_FILE} / "
            f"{RESEARCH_SHEET_NAME}."
        )

    return research


def get_research_for_fund(
    pru_access_name: str,
    research: dict[str, dict],
) -> tuple[dict | None, str]:
    """
    Match the master PruAccess name to Research Funds.xlsx.

    Matching is exact after whitespace/case normalization.

    Returns:

        (research_record, status)

    status:

        matched
        missing
        no_master_name
    """

    master_name = clean_text(
        pru_access_name
    )

    if not master_name:
        return None, "no_master_name"

    key = normalize_match_key(
        master_name
    )

    record = research.get(key)

    if record is None:
        return None, "missing"

    return record, "matched"


# =============================================================================
# HOLDINGS
# =============================================================================

def clean_holdings(
    holdings,
    label: str
) -> list[dict]:

    if (
        not isinstance(holdings, list)
        or not holdings
    ):
        raise ValueError(
            f"{label}: holdings list is empty "
            "or invalid."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"{label}: more than "
            f"{MAX_HOLDINGS} holdings."
        )

    cleaned = []

    for position, item in enumerate(
        holdings,
        start=1
    ):
        if not isinstance(item, dict):
            raise ValueError(
                f"{label}: holding {position} "
                "is not an object."
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
                f"{item.get('rank')} at "
                f"position {position}."
            )

        if not name:
            raise ValueError(
                f"{label}: holding {position} "
                "has no name."
            )

        if (
            isinstance(weight, bool)
            or not isinstance(
                weight,
                (int, float)
            )
            or not 0 <= weight <= 100
        ):
            raise ValueError(
                f"{label}: holding {position} "
                f"has invalid weight "
                f"{weight!r}."
            )

        cleaned.append(
            {
                "rank": position,
                "name": name,
                "weightPercent": weight,
                "weightText": clean_text(
                    item.get("weightText")
                ),
            }
        )

    return cleaned


def holdings_block(
    result: dict,
    stage: str,
    label: str
):
    """
    Return the topHoldings block.

    Returns None for an unresolved/failure result.
    """

    status = result.get(
        "status"
    )

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
            label
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


def resolve_holdings(
    excel_funds: dict[int, dict]
) -> dict[int, dict]:

    resolved: dict[int, dict] = {}

    def accept(
        result: dict,
        stage: str
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
                "is not in master Excel."
            )

        if row in resolved:
            return

        result_url = clean_text(
            result.get(
                "prudentialUrl"
            )
        )

        master_url = excel_funds[
            row
        ]["prudentialUrl"]

        if (
            result_url
            and result_url != master_url
        ):
            raise RuntimeError(
                f"{stage}: URL mismatch "
                f"for row {row}."
            )

        block = holdings_block(
            result,
            stage,
            f"{stage} row {row}"
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

    # ------------------------------------------------------------------
    # Baseline
    # ------------------------------------------------------------------

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
        []
    ):
        accept(
            result,
            "baseline"
        )

    # ------------------------------------------------------------------
    # Recovery 1 / 2 / 3
    # ------------------------------------------------------------------

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
                stage
            )

    return resolved


# =============================================================================
# PRUACCESS
# =============================================================================

def validate_bid_history(
    history: dict,
    label: str
) -> list[dict]:

    observations = history.get(
        "observations"
    )

    if (
        not isinstance(
            observations,
            list
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
            f"{label}: observationCount "
            "mismatch."
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
                str
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
                (int, float)
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
                "not chronological."
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


def load_pruaccess() -> dict[int, dict]:

    result = {}

    if not PRUACCESS_FUNDS_DIR.exists():

        print(
            f"WARNING: "
            f"{PRUACCESS_FUNDS_DIR} "
            "not found."
        )

        return result

    for directory in sorted(
        PRUACCESS_FUNDS_DIR.iterdir()
    ):

        match = re.match(
            r"^(\d+)_",
            directory.name
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


# =============================================================================
# FUND RESEARCH MERGE
# =============================================================================

def build_research_block(
    pru_access_name: str,
    research: dict[str, dict],
) -> tuple[dict, str]:

    record, status = get_research_for_fund(
        pru_access_name,
        research
    )

    if record is None:

        return (
            {
                "status": status,
                "source": str(
                    RESEARCH_EXCEL_FILE
                ),
                "worksheet": (
                    RESEARCH_SHEET_NAME
                ),
                "geographic1": None,
                "geographic2": None,
                "sector1": None,
                "sector2": None,
            },
            status,
        )

    return (
        {
            "status": "matched",
            "source": str(
                RESEARCH_EXCEL_FILE
            ),
            "worksheet": (
                RESEARCH_SHEET_NAME
            ),
            "researchFundName": (
                record[
                    "researchFundName"
                ]
            ),
            "geographic1": (
                record[
                    "geographic1"
                ]
                or None
            ),
            "geographic2": (
                record[
                    "geographic2"
                ]
                or None
            ),
            "sector1": (
                record[
                    "sector1"
                ]
                or None
            ),
            "sector2": (
                record[
                    "sector2"
                ]
                or None
            ),
        },
        "matched",
    )


# =============================================================================
# BUILD
# =============================================================================

def main() -> int:

    print("=" * 72)
    print(
        "VGRAT FMS - BUILD FUNDS DATA"
    )
    print("=" * 72)

    print(
        f"Master workbook: "
        f"{MASTER_EXCEL_FILE}"
    )

    print(
        f"Research workbook: "
        f"{RESEARCH_EXCEL_FILE}"
    )

    print(
        f"Research worksheet: "
        f"{RESEARCH_SHEET_NAME}"
    )

    print(
        f"Allow unresolved: "
        f"{ALLOW_UNRESOLVED}"
    )

    # ------------------------------------------------------------------
    # Load inputs
    # ------------------------------------------------------------------

    excel_funds = (
        read_excel_funds()
    )

    research_funds = (
        read_research_funds()
    )

    holdings = (
        resolve_holdings(
            excel_funds
        )
    )

    pruaccess = (
        load_pruaccess()
    )

    generated_at = (
        utc_now_iso()
    )

    # ------------------------------------------------------------------
    # Counters
    # ------------------------------------------------------------------

    fund_records = []
    bid_records = []

    gaps = []

    stage_counts: dict[str, int] = {}

    published = 0
    no_section = 0
    unresolved_holdings = 0

    bid_ok = 0
    bid_missing = 0
    total_observations = 0

    research_matched = 0
    research_missing = 0

    # ------------------------------------------------------------------
    # Build each master fund
    # ------------------------------------------------------------------

    for row in sorted(
        excel_funds
    ):

        excel = excel_funds[
            row
        ]

        pru = pruaccess.get(
            row
        )

        prudential = (
            (pru or {}).get(
                "prudential"
            )
            or {}
        )

        block = holdings.get(
            row
        )

        # ==============================================================
        # IDENTITY
        # ==============================================================

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
                or (
                    block or {}
                ).get(
                    "fundName"
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

        # ==============================================================
        # HOLDINGS
        # ==============================================================

        if block is None:

            unresolved_holdings += 1

            gaps.append(
                "Row "
                f"{row}: holdings "
                "unresolved after "
                "Recovery 3."
            )

            top_holdings = {
                "status": "unresolved",
                "source": None,
                "count": 0,
                "holdings": [],
            }

        else:

            top_holdings = {
                key: value
                for key, value
                in block.items()
                if key != "fundName"
            }

            source = block[
                "source"
            ]

            stage_counts[
                source
            ] = (
                stage_counts.get(
                    source,
                    0
                )
                + 1
            )

            if (
                block[
                    "status"
                ]
                == "published"
            ):
                published += 1
            else:
                no_section += 1

        # ==============================================================
        # FUND INFORMATION
        # ==============================================================

        fund_info = None

        if prudential:

            fund_info = {
                key: value
                for key, value
                in prudential.items()
                if key not in {
                    "raw",
                    "fundIdentifier",
                    "fundCode",
                    "fundName",
                }
            }

            # ----------------------------------------------------------
            # Dividend normalization
            # ----------------------------------------------------------

            dividend_rate = (
                fund_info.get(
                    "dividendRate"
                )
            )

            if (
                dividend_rate is not None
                and clean_text(
                    dividend_rate
                )
            ):
                fund_info[
                    "hasDividend"
                ] = True

            else:
                fund_info[
                    "hasDividend"
                ] = False

        # ==============================================================
        # RESEARCH
        # ==============================================================

        research_block, research_status = (
            build_research_block(
                excel[
                    "pruAccessName"
                ],
                research_funds,
            )
        )

        if research_status == "matched":

            research_matched += 1

        elif research_status == "missing":

            research_missing += 1

        # ==============================================================
        # FUNDS.JSON RECORD
        # ==============================================================

        fund_records.append(
            {
                **identity,

                "fund": fund_info,

                "research": research_block,

                "topHoldings": top_holdings,
            }
        )

        # ==============================================================
        # BID HISTORY
        # ==============================================================

        if pru is None:

            bid_missing += 1

            gaps.append(
                "Row "
                f"{row}: BID history "
                "missing."
            )

            bid_history = {
                "status": "unresolved",
                "priceType": "BID",
                "observationCount": 0,
                "observations": [],
            }

        else:

            history = (
                pru[
                    "bidHistory"
                ]
            )

            observations = (
                validate_bid_history(
                    history,
                    f"row {row} BID",
                )
            )

            bid_ok += 1

            total_observations += (
                len(
                    observations
                )
            )

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

                "observationCount": (
                    len(
                        observations
                    )
                ),

                "observations": (
                    observations
                ),
            }

        bid_records.append(
            {
                **identity,
                "bidHistory": bid_history,
            }
        )

    # ------------------------------------------------------------------
    # BUILD GATE
    # ------------------------------------------------------------------

    if gaps:

        print()
        print(
            "UNRESOLVED:"
        )

        for gap in gaps:
            print(
                f" - {gap}"
            )

        if not ALLOW_UNRESOLVED:

            print()
            print(
                "BUILD FAILED: "
                "unresolved core funds "
                "are present."
            )

            print(
                "Nothing was written."
            )

            print(
                "Set ALLOW_UNRESOLVED=1 "
                "to publish unresolved "
                "funds."
            )

            return 1

    # ------------------------------------------------------------------
    # Output envelope
    # ------------------------------------------------------------------

    envelope = {
        "schemaVersion": (
            SCHEMA_VERSION
        ),

        "generatedAtUtc": (
            generated_at
        ),

        "source": (
            str(
                MASTER_EXCEL_FILE
            )
        ),

        "fundCount": (
            len(
                fund_records
            )
        ),
    }

    # ------------------------------------------------------------------
    # funds.json
    # ------------------------------------------------------------------

    write_json_atomic(
        FUNDS_OUT,

        {
            **envelope,

            "summary": {
                "holdingsPublished": (
                    published
                ),

                "noHoldingsSection": (
                    no_section
                ),

                "holdingsUnresolved": (
                    unresolved_holdings
                ),

                "holdingsBySource": (
                    stage_counts
                ),

                "researchMatched": (
                    research_matched
                ),

                "researchMissing": (
                    research_missing
                ),

                "researchWorkbook": (
                    str(
                        RESEARCH_EXCEL_FILE
                    )
                ),

                "researchWorksheet": (
                    RESEARCH_SHEET_NAME
                ),
            },

            "funds": (
                fund_records
            ),
        },
    )

    # ------------------------------------------------------------------
    # bid_history.json
    # ------------------------------------------------------------------

    write_json_atomic(
        BID_OUT,

        {
            **envelope,

            "summary": {
                "fundsWithBidHistory": (
                    bid_ok
                ),

                "bidHistoryUnresolved": (
                    bid_missing
                ),

                "totalObservations": (
                    total_observations
                ),
            },

            "funds": (
                bid_records
            ),
        },

        compact=True,
    )

    # ------------------------------------------------------------------
    # Final summary
    # ------------------------------------------------------------------

    print()
    print(
        "BUILD COMPLETE"
    )

    print(
        f"Funds:                 "
        f"{len(fund_records)}"
    )

    print(
        f"Holdings published:    "
        f"{published}"
    )

    print(
        f"No holdings section:   "
        f"{no_section}"
    )

    print(
        f"Holdings unresolved:   "
        f"{unresolved_holdings}"
    )

    print(
        f"Holdings by source:    "
        f"{stage_counts}"
    )

    print(
        f"BID history funds:     "
        f"{bid_ok}"
    )

    print(
        f"BID observations:      "
        f"{total_observations}"
    )

    print(
        f"Research matched:      "
        f"{research_matched}"
    )

    print(
        f"Research missing:      "
        f"{research_missing}"
    )

    print(
        f"Wrote: {FUNDS_OUT}"
    )

    print(
        f"Wrote: {BID_OUT}"
    )

    return 0


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":

    try:
        raise SystemExit(
            main()
        )

    except Exception as error:

        print(
            f"\nFATAL BUILD ERROR: "
            f"{error}",
            file=sys.stderr,
        )

        raise SystemExit(1)
