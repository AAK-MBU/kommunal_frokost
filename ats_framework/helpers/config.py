"""Module for general configurations of the process"""

MAX_RETRY = 10

# ----------------------
# Queue population settings
# ----------------------
MAX_CONCURRENCY = 100  # tune based on backend capacity
MAX_RETRIES = 3  # transient failure retries per item
RETRY_BASE_DELAY = 0.5  # seconds (exponential backoff)

# ----------------------
# SharePoint settings
# ----------------------
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com"
SHAREPOINT_SITE_NAME = "MBURPA"
# SHAREPOINT_SITE_NAME = "Sundhed-Samarbejdsprojekter"
SHAREPOINT_DOCUMENT_LIBRARY = "Delte dokumenter"
SHAREPOINT_FOLDER_NAME = "Misc"

EXCEL_FILE_NAME = "test.xlsx"
EXCEL_SHEET_NAME = "Besvarelser"
MANGLER_SHEET_NAME = "Mangler"
