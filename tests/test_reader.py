from pathlib import Path

from excel_schema_engine import ExcelBuilder
from excel_schema_engine import ExcelReader


def test_reader(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    ws.append([1, "A", 1, 10])
    ws.append([2, "B", 2, 20])
    reader = ExcelReader(schema.sheets[0])

    rows = list(reader.iter_rows(ws))

    assert rows[0]["sku"] == 1
    assert rows[0]["delta"] == 10

    out = Path("tests/outputs/reader.xlsx")
    wb.save(out)