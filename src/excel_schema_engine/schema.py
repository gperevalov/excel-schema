from dataclasses import dataclass, field
from typing import List, Dict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

"""
Header object example:

schema = ExcelSchema(
    sheets=[
        SheetSchema(
            name="Basic settings",
            columns=[

                Column(
                    key="sku",
                    header="Product code"
                ),

                Column(
                    key="offer_id",
                    header="Seller code"
                ),

                Column(
                    key="strategy",
                    header="Strategy",
                    comment="Comment 12345",
                    subcolumns=[

                        Column(
                            key="price_delta_type",
                            header="Meaning"
                        ),

                        Column(
                            key="price_delta",
                            header="Price delta",
                            comment="different ..."
                        ),
                    ]
                ),
            ]
        )
    ]
)
"""


@dataclass
class Column:
    key: str # variable name
    header: str # header text
    comment: str | None = None # comment for cell
    style: str | None = None # cell style
    subcolumns: list["Column"] = field(default_factory=list) # list cell under header cell


@dataclass
class SheetSchema:
    name: str # sheet name
    columns: List[Column] # list columns


@dataclass
class CellStyle:
    font: Font | None = None # for openpyxl.styles.Font
    fill: PatternFill | None = None # for openpyxl.styles.PatternFill
    alignment: Alignment | None = None # for openpyxl.styles.Alignment
    border: Border | None = None # for openpyxl.styles.Border


@dataclass
class ExcelSchema:
    sheets: List[SheetSchema] # list sheets with columns
    styles: Dict[str, CellStyle] = field(default_factory=dict)
    author: str = '' # author for comments


@dataclass
class ExcelErrorsSchema:
    error_fill: CellStyle = field(default_factory=lambda: CellStyle(
        font=Font(bold=True, color='FFFFFF', size=11),
        fill=PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
        alignment=Alignment(horizontal='center', vertical='center'),
        border=Border(
            left=Side(border_style="thin", color="808080"),
            right=Side(border_style="thin", color="808080"),
            top=Side(border_style="thin", color="808080"),
            bottom=Side(border_style="thin", color="808080")
        )
    ))
    fixed_fill: CellStyle = field(default_factory=lambda:  CellStyle(
        font=Font(bold=True, color='000000', size=11),
        fill=PatternFill(start_color="78FF66", end_color="78FF66", fill_type="solid"),
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
        border=Border(
            left=Side(border_style="thin", color="808080"),
            right=Side(border_style="thin", color="808080"),
            top=Side(border_style="thin", color="808080"),
            bottom=Side(border_style="thin", color="808080")
        )
    ))
    author: str = ''
