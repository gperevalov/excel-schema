from typing import Iterator

from openpyxl.worksheet.worksheet import Worksheet

from excel_schema_engine import SheetSchema
from excel_schema_engine.utils import flatten_columns


class ExcelReader:

    def __init__(self, sheet_schema: SheetSchema):
        self.sheet_schema = sheet_schema
        self.column_map: dict[str, int] = self._build_column_map()

    def _build_column_map(self) -> dict[str, int]:

        return {
            column.key: idx + 1
            for idx, column in enumerate(
                flatten_columns(self.sheet_schema.columns)
            )
        }

    def parse_row(self, row: tuple) -> dict:

        return {
            key: row[col - 1]
            for key, col in self.column_map.items()
        }

    def iter_rows(
        self,
        ws: Worksheet,
        start_row: int = 3
    ) -> Iterator[dict]:

        for row in ws.iter_rows(min_row=start_row, values_only=True):
            yield self.parse_row(row)

    def read_all(self, ws: Worksheet, start_row: int = 3) -> list[dict]:

        return list(self.iter_rows(ws, start_row))