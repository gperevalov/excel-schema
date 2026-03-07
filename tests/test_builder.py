from pathlib import Path

from excel_schema_engine import ExcelBuilder


def test_build_excel(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    assert ws.cell(1, 1).value == "SKU"
    assert ws.cell(1, 2).value == "Offer"
    assert ws.cell(1, 3).value == "Strategy"
    assert ws.cell(2, 3).value == "Type"

    out = Path("tests/outputs/builder.xlsx")
    wb.save(out)