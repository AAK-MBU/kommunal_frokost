"""Module for handling application startup, and close"""

import logging
import os
from dataclasses import dataclass
from typing import Any

from dotenv import load_dotenv
from mbu_msoffice_integration.sharepoint_class import Sharepoint
from mbu_rpa_core.exceptions import ProcessError

from ats_framework.helpers import config

logger = logging.getLogger(__name__)


@dataclass
class AppContext:
    """Holds the running application(s) for use across the project."""

    app: Any | None = None


CONTEXT = AppContext()


def get_app() -> Any:
    """Return the running application instance."""
    if CONTEXT.app is None:
        raise ProcessError("Application not started - call startup() first")
    return CONTEXT.app


def startup():
    """Function for starting applications"""
    logger.info("Starting applications...")

    load_dotenv()

    sharepoint = Sharepoint(
        tenant=os.getenv("TENANT"),
        client_id=os.getenv("CLIENT_ID"),
        thumbprint=os.getenv("APPREG_THUMBPRINT"),
        cert_path=os.getenv("GRAPH_CERT_PEM"),
        site_url=config.SHAREPOINT_SITE_URL,
        site_name=config.SHAREPOINT_SITE_NAME,
        document_library=config.SHAREPOINT_DOCUMENT_LIBRARY,
    )

    # Sharepoint swallows authentication errors and leaves ctx as None
    if sharepoint.ctx is None:
        raise ProcessError(
            f"Could not authenticate to SharePoint site {config.SHAREPOINT_SITE_NAME}"
        )

    CONTEXT.app = sharepoint


def soft_close():
    """Function for closing applications softly"""
    logger.info("Closing applications softly...")

    # if CONTEXT.app is not None:
    #     CONTEXT.app.close_application()

    CONTEXT.app = None


def hard_close():
    """Function for closing applications hard"""
    logger.info("Closing applications hard...")

    # Kill the process here, e.g. os.system("taskkill /f /im SolteqTand.exe")

    CONTEXT.app = None


def close():
    """Function for closing applications softly or hardly if necessary"""
    try:
        soft_close()
    except Exception:
        logger.warning("Soft close failed, forcing hard close", exc_info=True)
        hard_close()


def reset():
    """Function for resetting application"""
    logger.info("Resetting applications...")
    close()
    startup()
