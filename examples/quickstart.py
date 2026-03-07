from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from excel_schema_engine import (
    ExcelSchema,
    SheetSchema,
    Column,
    CellStyle,
    ExcelBuilder,
    autosize_columns,
)
from excel_schema_engine.global_vars import Language


def build_schema() -> ExcelSchema:
    """Return a simple ExcelSchema used in the quickstart examples."""
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    styles = {
        "header": CellStyle(
            font=Font(bold=True, color="FFFFFF", size=11),
            fill=PatternFill("solid", fgColor="000066CC"),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=thin_border,
        ),
    }

    return ExcelSchema(
        author="Your Name",
        styles=styles,
        sheets=[
            SheetSchema(
                name="Products",
                columns=[
                    Column(key="sku", header="SKU", style="header"),
                    Column(key="offer", header="Offer", style="header"),
                    Column(
                        key="strategy",
                        header="Strategy",
                        style="header",
                        subcolumns=[
                            Column(key="delta_type", header="Type", style="header"),
                            Column(key="delta", header="Delta", style="header"),
                        ],
                    ),
                ],
            ),
        ],
    )


def main() -> None:
    """Generate an example Excel file with headers and sample rows."""
    schema = build_schema()

    # 1) Build an Excel template from the schema
    builder = ExcelBuilder(schema)
    wb = builder.build()
    ws = wb["Products"]

    # 2) Write data rows starting from row 3 (after headers)
    rows = [
        {"sku": 1, "offer": "A", "delta_type": 1, "delta": 10},
        {"sku": 2, "offer": "B", "delta_type": 2, "delta": 20},
    ]

    builder.write_rows(ws, "Products", rows, start_row=3, language=Language.EN)

    # 3) Optional: autosize columns
    autosize_columns(ws)

    wb.save("./examples/products.xlsx")


if __name__ == "__main__":
    main()

