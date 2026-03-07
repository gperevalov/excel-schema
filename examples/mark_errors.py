from openpyxl import load_workbook

from excel_schema_engine import ExcelErrorsSchema, ExcelErrors
from excel_schema_engine.global_vars import Language


def mark_errors(path: str = "./examples/products.xlsx") -> None:
    """Demonstrate how to mark and highlight errors directly in an Excel file."""
    wb = load_workbook(path)
    ws = wb["Products"]

    error_schema = ExcelErrorsSchema(author="Validator")
    errors = ExcelErrors(error_schema, language=Language.EN)

    # Mark a single cell as error
    errors.mark_error(
        ws,
        row=4,
        col=1,
        msg="SKU must be an integer",
    )

    # Highlight an entire row
    errors.highlight_row(ws, row=4)

    # Later, when the value is fixed
    errors.mark_fixed(
        ws,
        row=4,
        col=1,
        msg="Fixed by user",
    )

    wb.save("./examples/products_with_errors.xlsx")


if __name__ == "__main__":
    mark_errors()

