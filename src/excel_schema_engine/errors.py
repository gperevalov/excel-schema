from openpyxl.comments import Comment
from openpyxl.worksheet.worksheet import Worksheet


class ExcelErrors:

    def __init__(self, error_schema):
        self.error_schema = error_schema

    def mark_error(
        self,
        ws: Worksheet,
        row: int,
        col: int,
        msg: str | None = None,
        height: int = 79,
        width: int = 144
    ):
        """Mark cell as error"""
        cell = ws.cell(row=row, column=col)

        self._apply_style(cell, self.error_schema.error_fill)

        if msg is not None:
            cell.comment = Comment(
                msg,
                self.error_schema.author,
                width,
                height
            )

    def highlight_row(self, ws: Worksheet, row: int):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            self._apply_style(cell, self.error_schema.error_fill)

    def mark_fixed(
        self,
        ws: Worksheet,
        row: int,
        col: int,
        msg: str | None = None,
        height: int = 79,
        width: int = 144,
        check_msg: str | None = None,
    ):
        """Mark cell as fixed"""
        cell = ws.cell(row=row, column=col)

        if check_msg is not None:
            if cell.comment is None:
                return

            if check_msg.lower() not in cell.comment.text.lower():
                return

        self._apply_style(cell, self.error_schema.fixed_fill)

        if msg:
            cell.comment = Comment(
                msg,
                self.error_schema.author,
                width,
                height
            )

    @staticmethod
    def _apply_style(cell, style):

        if style.font:
            cell.font = style.font

        if style.fill:
            cell.fill = style.fill

        if style.alignment:
            cell.alignment = style.alignment

        if style.border:
            cell.border = style.border