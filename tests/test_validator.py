from pathlib import Path

from excel_schema_engine import ExcelBuilder
from excel_schema_engine import ExcelValidator


def test_validator(schema):

    builder = ExcelBuilder(schema)

    wb = builder.build()

    ws = wb["Products"]

    validator = ExcelValidator(schema.sheets[0])

    errors = validator.validate_headers(ws)

    assert errors == []

    out = Path("tests/outputs/validator.xlsx")
    wb.save(out)