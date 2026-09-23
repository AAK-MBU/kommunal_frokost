"""Module to handle item processing"""

import logging
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
from mbu_msoffice_integration.sharepoint_class import Sharepoint
from mbu_rpa_core.exceptions import ProcessError

from ats_framework.core.application_handler import get_app
from ats_framework.helpers import config

logger = logging.getLogger(__name__)

SID_COLUMN = "Besvarelses-ID"

COLUMNS = [
    "Dagtilbud",
    "Dagtilbudsleder",
    "LISID",
    "Enhedens kaldenavn",
    "Afdelingstype",
    "Ledertype",
    "Navn på leder",
    "Leder e-mail",
    "Ejertype",
    "Bestyrelsens beslutning (frokost/madpakke)",
    "Aldersopdelt valg ja/nej",
    "Type måltid (egenproduktion/samproduktion eller ekstern)",
    "Hvorfra samproduktion",
    "Ekstern leverandør",
    "Bemærkning",
    "Indsendt",
    SID_COLUMN,
]

FROKOST_ELLER_MADPAKKE_LABELS = {
    "frokost": "Frokostmåltid",
    "madpakke": "Madpakke",
}

ALDERSOPDELING_LABELS = {
    "aldersopdelt": "Ja",
    "samlet": "Nej",
}

FROKOSTTYPE_LABELS = {
    "egenproduktion": "Egenproduktion",
    "samproduktion_fra_en_anden_afdeling": "Samproduktion",
    "ekstern_leveret_fra_leverandoer_eller_skole": "Ekstern",
}


def format_timestamp(value: str) -> str:
    """Format an ISO timestamp as Danish local time, e.g. "21-09-2026 14:09"."""
    try:
        parsed = datetime.fromisoformat(value)
    except (TypeError, ValueError):
        return value or ""

    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(ZoneInfo("Europe/Copenhagen"))

    return parsed.strftime("%d-%m-%Y %H:%M")


def build_rows(submission: dict) -> list[dict]:
    """
    Turn one queued submission into Excel rows - one row per afdeling.

    Afdelinger found in masterdata but missing from the submission get a row
    as well, so it is visible that they have not been answered.
    """
    dagtilbud_md = submission.get("dagtilbud_masterdata", {})
    beslutning = FROKOST_ELLER_MADPAKKE_LABELS.get(
        submission.get("frokost_eller_madpakke", ""), ""
    )

    common = {
        "Dagtilbud": dagtilbud_md.get("enhedsnavn") or submission["dagtilbud_navn"],
        "Dagtilbudsleder": dagtilbud_md.get("leder_navn", ""),
        "Bestyrelsens beslutning (frokost/madpakke)": beslutning,
        "Indsendt": format_timestamp(
            submission.get("completed") or submission.get("form_submitted_date")
        ),
        SID_COLUMN: int(submission["sid"]),
    }

    def unit_columns(md: dict, fallback_name: str) -> dict:
        return {
            "LISID": md.get("lisid", ""),
            "Enhedens kaldenavn": md.get("enhedsnavn") or fallback_name,
            "Afdelingstype": md.get("afdelingstype", ""),
            "Ledertype": md.get("ledertype", ""),
            "Navn på leder": md.get("leder_navn", ""),
            "Leder e-mail": md.get("leder_email", ""),
            "Ejertype": md.get("ejertype", ""),
        }

    rows = []

    for afdeling in submission.get("afdelinger", []):
        rows.append(
            {
                **common,
                **unit_columns(
                    afdeling.get("masterdata", {}), afdeling["afdeling_navn"]
                ),
                "Aldersopdelt valg ja/nej": ALDERSOPDELING_LABELS.get(
                    afdeling["aldersopdeling"], ""
                ),
                "Type måltid (egenproduktion/samproduktion eller ekstern)": (
                    FROKOSTTYPE_LABELS.get(afdeling["frokosttype"], "")
                ),
                "Hvorfra samproduktion": afdeling["samproduktion_fra"],
                "Ekstern leverandør": afdeling["ekstern_leverandoer"],
                "Bemærkning": afdeling["bemaerkning"],
            }
        )

    for md in submission.get("ikke_besvarede_afdelinger", []):
        rows.append(
            {
                **common,
                **unit_columns(md, ""),
                "Bemærkning": "Ikke med i besvarelsen",
            }
        )

    # A madpakke answer has no afdelinger - still register the dagtilbud's decision
    if not rows:
        rows.append({**common, **unit_columns(dagtilbud_md, "")})

    return [{col: row.get(col, "") for col in COLUMNS} for row in rows]


def fetch_existing_sids(sharepoint: Sharepoint) -> set[int] | None:
    """
    Return the submission IDs already in the Excel file,
    or None if the file does not exist yet.
    """
    files = sharepoint.fetch_files_list(folder_name=config.SHAREPOINT_FOLDER_NAME)
    if files is None:
        raise ProcessError(
            f"Could not list files in SharePoint folder '{config.SHAREPOINT_FOLDER_NAME}'"
        )

    if config.EXCEL_FILE_NAME not in {f["Name"] for f in files}:
        return None

    content = sharepoint.fetch_file_using_open_binary(
        config.EXCEL_FILE_NAME, config.SHAREPOINT_FOLDER_NAME
    )
    if content is None:
        raise ProcessError(f"Could not download '{config.EXCEL_FILE_NAME}'")

    df = pd.read_excel(BytesIO(content), sheet_name=config.EXCEL_SHEET_NAME)

    return set(pd.to_numeric(df[SID_COLUMN], errors="coerce").dropna().astype(int))


def create_excel_file(sharepoint: Sharepoint, rows: list[dict]) -> None:
    """Create the Excel file with the given rows and upload it."""
    stream = BytesIO()
    pd.DataFrame(rows, columns=COLUMNS).to_excel(
        stream, index=False, engine="openpyxl", sheet_name=config.EXCEL_SHEET_NAME
    )

    sharepoint.upload_file_from_bytes(
        binary_content=stream.getvalue(),
        file_name=config.EXCEL_FILE_NAME,
        folder_name=config.SHAREPOINT_FOLDER_NAME,
    )


def process_item(item_data: dict, item_reference: str):
    """
    Write the rows for one submission to the Excel file in SharePoint.

    Creates the file if it does not exist, otherwise appends. Submissions
    already present in the file (matched on Besvarelses-ID) are skipped, so
    the process can safely be rerun.
    """
    sharepoint: Sharepoint = get_app()

    rows = build_rows(item_data)
    sid = int(item_data["sid"])

    existing_sids = fetch_existing_sids(sharepoint)

    if existing_sids is None:
        logger.info("Creating '%s' with submission %s", config.EXCEL_FILE_NAME, sid)
        create_excel_file(sharepoint, rows)

    elif sid in existing_sids:
        logger.info(
            "Submission %s (%s) already in '%s' - skipping",
            sid,
            item_reference,
            config.EXCEL_FILE_NAME,
        )
        return

    else:
        logger.info("Appending %d rows for submission %s", len(rows), sid)
        sharepoint.append_row_to_sharepoint_excel(
            required_headers=COLUMNS,
            folder_name=config.SHAREPOINT_FOLDER_NAME,
            excel_file_name=config.EXCEL_FILE_NAME,
            sheet_name=config.EXCEL_SHEET_NAME,
            new_rows=rows,
        )

    # Sort by Dagtilbud, then Enhedens kaldenavn
    sharepoint.format_and_sort_excel_file(
        folder_name=config.SHAREPOINT_FOLDER_NAME,
        excel_file_name=config.EXCEL_FILE_NAME,
        sheet_name=config.EXCEL_SHEET_NAME,
        sorting_keys=[
            {"key": "A", "ascending": True},
            {"key": "D", "ascending": True},
        ],
        bold_rows=[1],
        align_horizontal="left",
        align_vertical="top",
        column_widths=50,
        freeze_panes="A2",
    )

    # Sharepoint only prints upload errors, so verify the rows actually landed
    if sid not in (fetch_existing_sids(sharepoint) or set()):
        raise ProcessError(
            f"Submission {sid} was not found in '{config.EXCEL_FILE_NAME}' after upload"
        )
