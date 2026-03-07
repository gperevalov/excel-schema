from pathlib import Path

from excel_schema_engine import ExcelBuilder
from excel_schema_engine import ExcelErrors
from excel_schema_engine import ExcelErrorsSchema
from excel_schema_engine import autosize_columns

def test_errors(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    ws.append(["a", "this has been fixed", "b", "c"])
    ws.append([1, "A", 1, 10])
    autosize_columns(ws)
    errors = ExcelErrors(ExcelErrorsSchema())

    errors.mark_error(ws, 3, 2, "Invalid value")

    errors.mark_fixed(ws, 3, 2, "Fixed")

    errors.highlight_row(ws, 4)

    out = Path("tests/outputs/errors.xlsx")
    wb.save(out)