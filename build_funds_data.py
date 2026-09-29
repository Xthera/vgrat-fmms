#!/usr/bin/env python3

"""
VGrat FMS - BUILD FUNDS DATA

Merges the outputs of the extraction pipelines into two data files that
share the SAME layout (same envelope, same per-fund identity block):

    data/funds.json          fund info + Top Holdings
    data/bid_history.json    historical BID observations

MASTER UNIVERSE
===============

Funds Links.xlsm, Column A (Prudential URL) / Column B (PruAccess name).

Every populated Excel row is evaluated independently.

PER-FUND PUBLICATION RULE
=========================

Each fund is processed independently.

1. SUCCESS + previous record exists
       -> replace the previous record with the newly extracted data.

2. SUCCESS + no previous record
       -> insert the newly extracted fund.

3. FAILURE + previous record exists
       -> retain the previous record unchanged.

4. FAILURE + no previous record
       -> do NOT publish a record for that fund.

A failure for one fund NEVER prevents successful funds from being
published.

DATA RETRIEVAL DATE
===================

Every newly successful fund receives:

    "dataRetrievedDate": "YYYY-MM-DD"

using the Singapore calendar date.

A retained fund keeps its previous dataRetrievedDate because that is
the date on which the retained data was actually retrieved successfully.

The envelope also contains:

    generatedAtUtc

which records when this complete build was executed.

HOLDINGS RESOLUTION
===================

Stage order:

    baseline -> recovery1 -> recovery2 -> recovery3

- A fund that succeeded in the baseline keeps its baseline result.
- A fund the baseline marked "no_holdings_section" is kept AS IS.
- A fund that FAILED in the baseline is replaced by the first recovery
  stage that resolved it.
- A fund still failed after Recovery 3 has unresolved holdings.

BID HISTORY
===========

Read from:

    output_pruaccess/funds/<row>_<id>/

using:

    bid_history.json
    prudential_fund.json

DIVIDEND NORMALIZATION
======================

Final funds.json rule:

    - If dividendRate is non-empty -> hasDividend = true
    - If dividendRate is empty/missing -> hasDividend = false

The existing dividendRate value is never modified.

IMPORTANT
=========

Previous published data is read from:

    data/funds.json
    data/bid_history.json

The previous files are NEVER used as the source of newly retrieved
fund information. They are used ONLY as fallback when the current
fund cannot be successfully built.

This means:

    successful current data always wins
    previous data is fallback only
"""


from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
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

SINGAPORE_TZ = ZoneInfo("Asia/Singapore")

# Retained for compatibility with existing workflow/environment.
# The new per-fund publication model does not require the entire build
# to fail when unresolved funds exist.
ALLOW_UNRESOLVED = os.environ.get(
    "ALLOW_UNRESOLVED", ""
).strip().lower() in {"1", "true", "yes"}


# =============================================================================
# HELPERS
# =============================================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\xa0", " "),
    ).strip()


def utc_now_iso() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
    )


def singapore_today() -> str:
    """
    Return the current Singapore calendar date.

    Example:

        2026-09-30
    """

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


def load_previous_file(path: Path) -> dict:
    """
    Load an existing published output.

    Missing previous files are treated as an empty publication rather
    than as an error.

    This is important for first-time insertion of successfully retrieved
    funds.
    """

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

    funds = data.get("funds")

    if not isinstance(funds, list):
        raise RuntimeError(
            f"Previous published file has no valid funds list: {path}"
        )

    return data


def index_previous_records(
    data: dict,
) -> dict[int, dict]:
    """
    Index previous records by Excel row.

    Excel row is the stable controlling identity in this project.
    """

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


# =============================================================================
# DIVIDEND NORMALIZATION
# =============================================================================

def normalize_dividend_fields(
    fund_info: dict,
) -> dict:
    """
    Ensure hasDividend is always consistent with dividendRate.

    Rule:

        non-empty dividendRate -> True
        empty/missing dividendRate -> False

    The existing dividendRate value itself is preserved unchanged.
    """

    dividend_rate = clean_text(
        fund_info.get("dividendRate")
    )

    fund_info["hasDividend"] = bool(
        dividend_rate
    )

    return fund_info


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> dict[int, dict]:

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


# =============================================================================
# HOLDINGS
# =============================================================================

def clean_holdings(
    holdings,
    label: str,
) -> list[dict]:

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


def holdings_block(
    result: dict,
    stage: str,
    label: str,
):

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


def resolve_holdings(
    excel_funds: dict[int, dict],
) -> dict[int, dict]:

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

        # A resolved fund is never replaced by
        # a later stage.
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

    # ---------------------------------------------------------------------
    # BASELINE
    # ---------------------------------------------------------------------

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

    # ---------------------------------------------------------------------
    # RECOVERY STAGES
    # ---------------------------------------------------------------------

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


# =============================================================================
# PRUACCESS
# =============================================================================

def validate_bid_history(
    history: dict,
    label: str,
) -> list[dict]:

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


def load_pruaccess() -> dict[int, dict]:

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


# =============================================================================
# PREVIOUS DATA HELPERS
# =============================================================================

def previous_fund_record(
    previous_funds: dict[int, dict],
    row: int,
) -> dict | None:

    record = previous_funds.get(row)

    if not isinstance(
        record,
        dict,
    ):
        return None

    return record


def previous_bid_record(
    previous_bids: dict[int, dict],
    row: int,
) -> dict | None:

    record = previous_bids.get(row)

    if not isinstance(
        record,
        dict,
    ):
        return None

    return record


def make_retained_fund_record(
    previous: dict,
) -> dict:

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


def make_retained_bid_record(
    previous: dict,
) -> dict:

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


# =============================================================================
# BUILD
# =============================================================================

def main() -> int:

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

    # ---------------------------------------------------------------------
    # MASTER INPUTS
    # ---------------------------------------------------------------------

    excel_funds = read_excel_funds()

    holdings = resolve_holdings(
        excel_funds
    )

    pruaccess = load_pruaccess()

    # ---------------------------------------------------------------------
    # PREVIOUS PUBLISHED DATA
    # ---------------------------------------------------------------------

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

    # ---------------------------------------------------------------------
    # OUTPUT COLLECTIONS
    # ---------------------------------------------------------------------

    fund_records = []
    bid_records = []

    # ---------------------------------------------------------------------
    # COUNTERS
    # ---------------------------------------------------------------------

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

    # ---------------------------------------------------------------------
    # PROCESS EVERY EXCEL FUND INDEPENDENTLY
    # ---------------------------------------------------------------------

    for row in sorted(excel_funds):

        excel = excel_funds[row]

        pru = pruaccess.get(row)

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

        # ================================================================
        # IDENTITY
        # ================================================================

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

        # ================================================================
        # CURRENT HOLDINGS
        # ================================================================

        block = holdings.get(row)

        current_holdings_success = (
            block is not None
        )

        # ================================================================
        # CURRENT FUND INFORMATION
        # ================================================================

        current_fund_success = bool(
            prudential
        )

        # ================================================================
        # CURRENT BID HISTORY
        # ================================================================

        current_bid_success = False
        current_bid_history = None
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

        # ================================================================
        # DETERMINE COMPLETE FUND SUCCESS
        # ================================================================
        #
        # A newly published fund requires:
        #
        #   1. Prudential fund data
        #   2. resolved holdings
        #   3. valid BID history
        #
        # "no_holdings_section" counts as a successful holdings result.
        #
        # ================================================================

        complete_success = (
            current_fund_success
            and current_holdings_success
            and current_bid_success
        )

        # ================================================================
        # SUCCESSFUL FUND
        # ================================================================

        if complete_success:

            # ------------------------------------------------------------
            # FUND INFO
            # ------------------------------------------------------------

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

            # ------------------------------------------------------------
            # HOLDINGS
            # ------------------------------------------------------------

            top_holdings = {
                k: v
                for k, v in block.items()
                if k != "fundName"
            }

            stage = block["source"]

            stage_counts[stage] = (
                stage_counts.get(
                    stage,
                    0,
                )
                + 1
            )

            if (
                block["status"]
                == "published"
            ):
                holdings_published += 1

            else:
                holdings_no_section += 1

            # ------------------------------------------------------------
            # DATA STATUS
            # ------------------------------------------------------------

            if previous_fund is None:
                data_status = "inserted"
                fund_inserted += 1

            else:
                data_status = "updated"
                fund_updated += 1

            # ------------------------------------------------------------
            # FINAL FUND RECORD
            # ------------------------------------------------------------

            fund_record = {
                **identity,

                "dataStatus": data_status,

                "dataRetrievedDate": (
                    retrieval_date
                ),

                "fund": fund_info,

                "topHoldings": top_holdings,
            }

            fund_records.append(
                fund_record
            )

            # ------------------------------------------------------------
            # BID
            # ------------------------------------------------------------

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
                f"({data_status})"
            )

            continue

        # ================================================================
        # FAILED CURRENT FUND
        # ================================================================

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

        # ================================================================
        # PREVIOUS FUND EXISTS -> RETAIN
        # ================================================================

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

        # ================================================================
        # BID RETENTION IS INDEPENDENT
        # ================================================================

        if current_bid_success:

            # This branch is normally reached only when the complete
            # fund failed because of holdings or fund metadata.
            #
            # A valid current BID should still replace/insert the BID
            # record independently.

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

                    "bidHistory": (
                        bid_history
                    ),
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

        "generatedAtUtc": (
            generated_at
        ),

        "dataRetrievedDate": (
            retrieval_date
        ),

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
        },

        "funds": fund_records,
    }

    bid_envelope = {
        "schemaVersion": SCHEMA_VERSION,

        "generatedAtUtc": (
            generated_at
        ),

        "dataRetrievedDate": (
            retrieval_date
        ),

        "source": str(
            EXCEL_FILE
        ),

        "fundCount": len(
            bid_records
        ),

        "summary": {
            "fundsUpdated": (
                bid_updated
            ),

            "fundsInserted": (
                bid_inserted
            ),

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
    #
    # Important:
    #
    # The output is now written even when individual funds fail.
    #
    # Only funds with no successful current data AND no previous record
    # are omitted.
    #
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


if __name__ == "__main__":

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
