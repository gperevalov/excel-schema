from pathlib import Path

from excel_schema_engine import ExcelBuilder
from excel_schema_engine import autosize_columns
from excel_schema_engine.global_vars import Language


def test_errors(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    rows = [
        {
            "sku": "a",
            "offer": "this has been fixed",
            "delta_type": "b",
            "delta": "c"
        },
        {
            "sku": 1,
            "offer": "A",
            "delta_type": 1,
            "delta": 10
        },
        {
            "sku": 3,
            "offer": "B",
            "delta_type": 2,
            "delta": 16
        }
    ]

    builder.write_rows(ws, "Products", rows, 3, Language.BOBR)
    autosize_columns(ws)

    out = Path("tests/outputs/write_rows.xlsx")
    wb.save(out)