"""Module to hande queue population"""

import asyncio
import json
import logging
from datetime import datetime

from automation_server_client import Workqueue
from mbu_rpa_core.database import RPAConnection

from ats_framework.helpers import config

logger = logging.getLogger(__name__)

AFDELING_FIELDS = {
    "afdeling_navn": "afdeling",
    "aldersopdeling": "afdelings_specifikt_aldersopdeling",
    "frokosttype": "hvorfra_skal_der_bestilles_mad",
    "samproduktion_fra": "fra_hvilken_afdeling",
    "ekstern_leverandoer": "fra_hvilken_ekstern_levering",
    "bemaerkning": "saerlige_bemaerkning",
}


def _clean(value) -> str:
    """Return value as a stripped string, treating None as empty."""
    return str(value).strip() if value is not None else ""


def _to_json_safe(value):
    """Convert database values that are not JSON serializable."""
    if isinstance(value, datetime):
        return value.isoformat()
    return value


def parse_dagtilbud(vaelg_dagtilbud: str) -> dict:
    """
    Split the selected dagtilbud value.

    Format: "<dagtilbud_id>--<navn>--<antal_afdelinger>--<flag>",
    e.g. "872389--Børnegården Rundhøj (SDT)--1--1".
    """
    parts = [p.strip() for p in vaelg_dagtilbud.split("--")] if vaelg_dagtilbud else []
    parts += [""] * (4 - len(parts))

    return {
        "dagtilbud_id": parts[0],
        "dagtilbud_navn": parts[1],
        "antal_afdelinger": int(parts[2]) if parts[2].isdigit() else None,
        "dagtilbud_flag": parts[3],
    }


def parse_afdeling(row: dict) -> dict:
    """Flatten a single afdeling row from afdelinger_v1 / afdelinger_v2."""
    afdeling = {key: _clean(row.get(field)) for key, field in AFDELING_FIELDS.items()}

    # The afdeling select may deliver "<id>||<navn>" - split it if so
    afdeling["afdeling_id"] = ""
    if "||" in afdeling["afdeling_navn"]:
        afdeling_id, _, navn = afdeling["afdeling_navn"].partition("||")
        afdeling["afdeling_id"] = afdeling_id.strip()
        afdeling["afdeling_navn"] = navn.strip()

    return afdeling


def parse_submission(form_data: dict) -> dict:
    """
    Extract the relevant information from a kommunal_frokost submission.

    Dagtilbud with 1 afdeling (twig_page2_content == "1") answer frokost/madpakke
    directly and fill afdelinger_v1. Dagtilbud with several afdelinger
    (twig_page2_content == "2") fill afdelinger_v2, which also holds the
    aldersopdeling choice per afdeling.
    """
    data = form_data.get("data", {})
    entity = form_data.get("entity", {})

    use_v2 = _clean(data.get("twig_page2_content")) == "2"
    rows = data.get("afdelinger_v2" if use_v2 else "afdelinger_v1") or []

    afdelinger = [parse_afdeling(r) for r in rows if _clean(r.get("afdeling"))]

    return {
        "sid": _clean((entity.get("sid") or [{}])[0].get("value")),
        "completed": _clean((entity.get("completed") or [{}])[0].get("value")),
        **parse_dagtilbud(_clean(data.get("vaelg_dagtilbud"))),
        "flere_afdelinger": use_v2,
        "frokost_eller_madpakke": _clean(data.get("frokost_eller_madpakke")),
        "afdelinger": afdelinger,
    }


MASTERDATA_SQL = """
    SELECT
        [LISID],
        [ENHNAVN],
        [KALDENAVN_KORT],
        [LOS_NAVN],
        [LOSID],
        [ORG_REFERENCE_TIL],
        [LEDERNAVN],
        [E_MAIL],
        [LTYPE_TXT],
        [AFDTYPE_TXT],
        [EJER_TXT]
    FROM
        [BuMasterdata].[dbo].[VIEW_MD_STAMDATA_AKTUEL]
    WHERE
        LOSID = ?
        OR (
            -- Same afdeling filter as get_dagtilbud_afdelinger in the formular
            ORG_REFERENCE_TIL = ?
            AND HOMR = 3
            AND AFDTYPE != 3
        )
"""


def _normalize_name(name: str) -> str:
    """Normalize a unit name for matching (case and whitespace insensitive)."""
    return " ".join(_clean(name).casefold().split())


def _masterdata_info(row: dict | None) -> dict:
    """Pick the masterdata fields used in the report."""
    row = row or {}
    return {
        "lisid": _clean(row.get("LISID")),
        "losid": _clean(row.get("LOSID")),
        "enhedsnavn": _clean(row.get("ENHNAVN")),
        "leder_navn": _clean(row.get("LEDERNAVN")),
        "leder_email": _clean(row.get("E_MAIL")),
        "ledertype": _clean(row.get("LTYPE_TXT")),
        "afdelingstype": _clean(row.get("AFDTYPE_TXT")),
        "ejertype": _clean(row.get("EJER_TXT")),
    }


def fetch_masterdata(rpa_conn: RPAConnection, dagtilbud_id: str) -> tuple[dict, list]:
    """
    Fetch the dagtilbud itself (LOSID = dagtilbud_id) and its afdelinger
    (ORG_REFERENCE_TIL = dagtilbud_id) from BuMasterdata.

    Returns:
        (dagtilbud_row, afdeling_rows)
    """
    rows = (
        rpa_conn.execute_query(
            MASTERDATA_SQL, params=[dagtilbud_id, dagtilbud_id], return_dict=True
        )
        or []
    )

    dagtilbud_row = next(
        (r for r in rows if _clean(r.get("LOSID")) == dagtilbud_id), None
    )
    afdeling_rows = [r for r in rows if _clean(r.get("LOSID")) != dagtilbud_id]

    return dagtilbud_row, afdeling_rows


def enrich_with_masterdata(
    submission: dict, dagtilbud_row: dict | None, afdeling_rows: list
) -> None:
    """
    Add dagtilbudsleder and leder per afdeling to the submission (in place).

    Afdelinger from the form are matched by name against ENHNAVN, LOS_NAVN and
    KALDENAVN_KORT. The dagtilbud row itself is included as a candidate, since
    dagtilbud with a single afdeling may not have any child units.
    """
    if dagtilbud_row is None:
        logger.warning(
            "No masterdata found for dagtilbud %s (%s)",
            submission["dagtilbud_id"],
            submission["dagtilbud_navn"],
        )

    submission["dagtilbud_masterdata"] = _masterdata_info(dagtilbud_row)

    candidates = afdeling_rows + ([dagtilbud_row] if dagtilbud_row else [])
    by_name: dict[str, dict] = {}
    for row in candidates:
        for field in ("ENHNAVN", "LOS_NAVN", "KALDENAVN_KORT"):
            key = _normalize_name(row.get(field))
            if key:
                by_name.setdefault(key, row)

    matched_lisids = set()
    for afdeling in submission["afdelinger"]:
        row = None
        if afdeling["afdeling_id"]:
            row = next(
                (
                    r
                    for r in candidates
                    if afdeling["afdeling_id"]
                    in (_clean(r.get("LISID")), _clean(r.get("LOSID")))
                ),
                None,
            )
        row = row or by_name.get(_normalize_name(afdeling["afdeling_navn"]))

        if row is None:
            logger.warning(
                "No masterdata match for afdeling '%s' in dagtilbud %s",
                afdeling["afdeling_navn"],
                submission["dagtilbud_id"],
            )
        else:
            matched_lisids.add(_clean(row.get("LISID")))

        afdeling["masterdata"] = _masterdata_info(row)

    # Afdelinger registered in masterdata that were not part of the submission
    submission["ikke_besvarede_afdelinger"] = [
        _masterdata_info(r)
        for r in afdeling_rows
        if _clean(r.get("LISID")) not in matched_lisids
    ]


SUBMISSIONS_SQL = """
    SELECT
        [form_id],
        [form_sid],
        [form_type],
        [form_source],
        CONVERT(nvarchar(33), [form_submitted_date], 127) AS [form_submitted_date],
        [destination_system],
        [status],
        [response],
        CONVERT(nvarchar(33), [documented_date], 127) AS [documented_date],
        [form_data],
        CONVERT(nvarchar(33), [last_time_modified], 127) AS [last_time_modified]
    FROM
        [RPA].[journalizing].[view_Journalizing] AS vj
    WHERE
        form_type = 'kommunal_frokost'
    ORDER BY
        vj.[form_submitted_date] DESC
"""

# Same selection of dagtilbud as get_dagtilbud in the formular
ALL_DAGTILBUD_SQL = """
    SELECT
        d.LOSID,
        d.DAGTBNR_TXT,
        d.LEDERNAVN,
        d.E_MAIL,
        COUNT(DISTINCT a.LOSID) AS antal_afdelinger,
        CASE
            WHEN d.EJERTYPE = 2 THEN 1
            ELSE 0
        END AS sdt
    FROM
        [BuMasterdata].[dbo].[VIEW_MD_STAMDATA_AKTUEL] d
    LEFT JOIN
        [BuMasterdata].[dbo].[VIEW_MD_STAMDATA_AKTUEL] a
        ON a.ORG_REFERENCE_TIL = d.LOSID
        AND a.AFDTYPE != 3
        AND a.HOMR = 3
    WHERE
        d.HOMR = 3
        AND d.AFDTYPE = 1
    GROUP BY
        d.LOSID,
        d.DAGTBNR_TXT,
        d.LEDERNAVN,
        d.E_MAIL,
        d.EJERTYPE
    ORDER BY
        d.DAGTBNR_TXT
"""


def fetch_submissions(rpa_conn: RPAConnection) -> list[dict]:
    """Fetch all kommunal_frokost submissions, with form_data parsed from JSON."""
    rows = rpa_conn.execute_query(SUBMISSIONS_SQL, return_dict=True) or []

    for row in rows:
        if isinstance(row.get("form_data"), str):
            row["form_data"] = json.loads(row["form_data"])

    return rows


def fetch_all_dagtilbud(rpa_conn: RPAConnection) -> list[dict]:
    """Fetch every dagtilbud that can be selected in the formular."""
    return rpa_conn.execute_query(ALL_DAGTILBUD_SQL, return_dict=True) or []


def retrieve_items_for_queue() -> list[dict]:
    """Function to populate queue"""
    data = []
    references = []

    rpa_conn = RPAConnection(db_env="PROD", commit=False)
    with rpa_conn:
        rows = fetch_submissions(rpa_conn)

        logger.info("Fetched %d kommunal_frokost submissions", len(rows))

        masterdata_cache: dict[str, tuple[dict, list]] = {}

        for row in rows:
            submission = parse_submission(row.get("form_data") or {})
            submission["form_id"] = _clean(row.get("form_id"))
            submission["form_submitted_date"] = _to_json_safe(
                row.get("form_submitted_date")
            )

            dagtilbud_id = submission["dagtilbud_id"]
            if dagtilbud_id not in masterdata_cache:
                masterdata_cache[dagtilbud_id] = fetch_masterdata(
                    rpa_conn, dagtilbud_id
                )
            enrich_with_masterdata(submission, *masterdata_cache[dagtilbud_id])

            references.append(submission["form_id"])
            data.append(submission)

    items = [
        {"reference": ref, "data": d} for ref, d in zip(references, data, strict=True)
    ]

    return items


def create_sort_key(item: dict) -> str:
    """
    Create a sort key based on the entire JSON structure.
    Converts the item to a sorted JSON string for consistent ordering.
    """
    return json.dumps(item, sort_keys=True, ensure_ascii=False)


async def concurrent_add(workqueue: Workqueue, items: list[dict]) -> None:
    """
    Populate the workqueue with items to be processed.
    Uses concurrency and retries with exponential backoff.

    Args:
        workqueue (Workqueue): The workqueue to populate.
        items (list[dict]): List of items to add to the queue.

    Returns:
        None

    Raises:
        Exception: If adding an item fails after all retries.
    """
    sem = asyncio.Semaphore(config.MAX_CONCURRENCY)

    async def add_one(it: dict):
        reference = str(it.get("reference") or "")
        data = it

        async with sem:
            for attempt in range(1, config.MAX_RETRIES + 1):
                try:
                    await asyncio.to_thread(workqueue.add_item, data, reference)
                    logger.info("Added item to queue with reference: %s", reference)
                    return True

                except Exception as e:  # noqa: BLE001 - retry on any failure
                    if attempt >= config.MAX_RETRIES:
                        logger.error(
                            "Failed to add item %s after %d attempts: %s",
                            reference,
                            attempt,
                            e,
                        )
                        return False

                    backoff = config.RETRY_BASE_DELAY * (2 ** (attempt - 1))

                    logger.warning(
                        "Error adding %s (attempt %d/%d). Retrying in %.2fs... %s",
                        reference,
                        attempt,
                        config.MAX_RETRIES,
                        backoff,
                        e,
                    )
                    await asyncio.sleep(backoff)

    if not items:
        logger.info("No new items to add.")
        return

    sorted_items = sorted(items, key=create_sort_key)
    logger.info(
        "Processing %d items sorted by complete JSON structure", len(sorted_items)
    )

    results = await asyncio.gather(*(add_one(i) for i in sorted_items))
    successes = sum(1 for r in results if r)
    failures = len(results) - successes

    logger.info(
        "Summary: %d succeeded, %d failed out of %d", successes, failures, len(results)
    )
