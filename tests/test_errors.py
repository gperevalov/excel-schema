from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from excel_schema_engine import ExcelBuilder, CellStyle
from excel_schema_engine import ExcelErrors
from excel_schema_engine import ExcelErrorsSchema
from excel_schema_engine import autosize_columns
from excel_schema_engine.global_vars import Language


def test_errors(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    ws.append(["a", "this has been fixed", "b", "c"])
    ws.append([1, "A", 4, 10])
    ws.append([2, "B", 3, 20])
    ws.append([3, "C", 2, 30])
    ws.append([4, "D", 1, 40])
    autosize_columns(ws)
    errors = ExcelErrors(ExcelErrorsSchema(), Language.BOBR)

    errors.mark_error(ws, 3, 2, "Invalid value")

    errors.mark_error(
        ws,
        5,
        4,
        'misstake',
        custom_fill=CellStyle(
            font=Font(bold=True, color='000000', size=11),
            fill=PatternFill(start_color="78FF66", end_color="78FF66", fill_type="solid"),
            alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
            border=Border(
                left=Side(border_style="thin", color="808080"),
                right=Side(border_style="thin", color="808080"),
                top=Side(border_style="thin", color="808080"),
                bottom=Side(border_style="thin", color="808080")
            )
        )
    )

    errors.mark_fixed(ws, 3, 2, "Fixed")


    errors.highlight_row(ws, 4)

    out = Path("tests/outputs/errors.xlsx")
    wb.save(out)