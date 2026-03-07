from typing import List
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_schema_engine import Column


def autosize_columns(ws: Worksheet):
    """Resize worksheet columns to fit the longest cell value in each column."""
    for col in ws.columns:

        max_len = max(
            len(str(cell.value)) if cell.value else 0
            for cell in col
        )

        ws.column_dimensions[
            get_column_letter(col[0].column)
        ].width = max_len + 2


def flatten_columns(columns: List[Column]) -> List[Column]:
    """Flatten a list of (possibly nested) Column objects into a flat list."""
    result: List[Column] = []

    for column in columns:

        if column.subcolumns:
            result.extend(column.subcolumns)
        else:
            result.append(column)

    return result
