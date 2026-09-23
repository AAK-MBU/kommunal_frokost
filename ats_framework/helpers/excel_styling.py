"""Styling of the report sheet, mirroring "Aldersopdelt og type 2024" in helpers"""

import math
import textwrap

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

FONT_NAME = "Arial"

# Header colours per column group (theme colours from the example resolved to RGB)
MASTERDATA_FILL = "C4BD97"
BESVARELSE_FILL = "E6B9B8"
META_FILL = "B7DEE8"

# Data rows alternate between these per dagtilbud, so each dagtilbud stands out
BAND_FILLS = ("FFFFFF", "D0D0D0")

# Column name -> (header fill, column width)
COLUMN_STYLES = {
    "Dagtilbud": (MASTERDATA_FILL, 18),
    "Dagtilbudsleder": (MASTERDATA_FILL, 22),
    "LISID": (MASTERDATA_FILL, 8),
    "Enhedens kaldenavn": (MASTERDATA_FILL, 28),
    "Afdelingstype": (MASTERDATA_FILL, 17),
    "Ledertype": (MASTERDATA_FILL, 20),
    "Navn på leder": (MASTERDATA_FILL, 22),
    "Leder e-mail": (MASTERDATA_FILL, 26),
    "Ejertype": (MASTERDATA_FILL, 11),
    "Bestyrelsens beslutning (frokost/madpakke)": (BESVARELSE_FILL, 20),
    "Aldersopdelt valg ja/nej": (BESVARELSE_FILL, 14),
    "Type måltid (egenproduktion/samproduktion eller ekstern)": (BESVARELSE_FILL, 26),
    "Hvorfra samproduktion": (BESVARELSE_FILL, 24),
    "Ekstern leverandør": (BESVARELSE_FILL, 24),
    "Bemærkning": (BESVARELSE_FILL, 34),
    "Indsendt": (META_FILL, 16),
    "Besvarelses-ID": (META_FILL, 13),
}

# Font sizes, and estimates used for row heights:
# characters per unit of column width, and line height in points
HEADER_FONT_SIZE = 9
DATA_FONT_SIZE = 8
CHARS_PER_WIDTH_UNIT = {HEADER_FONT_SIZE: 1.0, DATA_FONT_SIZE: 1.2}
LINE_HEIGHT = {HEADER_FONT_SIZE: 12.0, DATA_FONT_SIZE: 11.25}
ROW_PADDING = 4  # extra points added to the calculated height
MIN_ROW_HEIGHT = 15

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _fill(rgb: str) -> PatternFill:
    return PatternFill(fill_type="solid", start_color=rgb, end_color=rgb)


def _line_count(value, width: float, font_size: int) -> int:
    """Estimate the number of lines a wrapped cell value takes up."""
    if value is None or value == "":
        return 1

    chars_per_line = max(1, int(width * CHARS_PER_WIDTH_UNIT[font_size]))

    return sum(
        max(1, len(textwrap.wrap(line, chars_per_line, break_long_words=True)))
        for line in str(value).split("\n")
    )


def _row_height(cells, widths: dict, font_size: int) -> float:
    """Height in points fitting the tallest cell in the row, plus padding."""
    lines = max(
        (_line_count(c.value, widths[c.column_letter], font_size) for c in cells),
        default=1,
    )
    return max(MIN_ROW_HEIGHT, math.ceil(lines * LINE_HEIGHT[font_size] + ROW_PADDING))


def style_report_sheet(ws: Worksheet) -> None:
    """
    Apply the report styling to a sheet with a single header row.

    Has to run after format_and_sort_excel_file, since that rewrites all rows
    and resets fonts and alignment.
    """
    headers = [cell.value for cell in ws[1]]

    # Rows are sorted by dagtilbud - switch band colour whenever it changes
    dagtilbud_col = headers.index("Dagtilbud") + 1
    row_fills = {}
    band, previous = 1, object()
    for row_idx in range(2, ws.max_row + 1):
        dagtilbud = ws.cell(row=row_idx, column=dagtilbud_col).value
        if dagtilbud != previous:
            band, previous = 1 - band, dagtilbud
        row_fills[row_idx] = _fill(BAND_FILLS[band])

    for col_idx, header in enumerate(headers, start=1):
        header_fill, width = COLUMN_STYLES.get(header, (MASTERDATA_FILL, 15))
        is_dagtilbud = header == "Dagtilbud"

        column = ws.iter_rows(min_col=col_idx, max_col=col_idx)
        for (cell,) in column:
            cell.border = BORDER

            if cell.row == 1:
                cell.fill = _fill(header_fill)
                cell.font = Font(name=FONT_NAME, size=HEADER_FONT_SIZE, bold=True)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            else:
                cell.font = Font(name=FONT_NAME, size=DATA_FONT_SIZE, bold=is_dagtilbud)
                cell.alignment = Alignment(
                    horizontal="left", vertical="top", wrap_text=True
                )
                cell.fill = row_fills[cell.row]

        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    # Excel does not auto fit rows in generated files, so estimate the heights
    widths = {
        letter: ws.column_dimensions[letter].width
        for letter in (c.column_letter for c in ws[1])
    }
    for row in ws.iter_rows():
        font_size = HEADER_FONT_SIZE if row[0].row == 1 else DATA_FONT_SIZE
        ws.row_dimensions[row[0].row].height = _row_height(row, widths, font_size)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
