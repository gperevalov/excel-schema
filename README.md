## Excel Schema Engine

A lightweight Python library for **generating**, **validating**, and **reading** Excel files using a **declarative schema**.

Define your Excel structure once and use it to:

- generate Excel templates
- validate uploaded files
- read rows as structured data
- highlight errors directly in Excel

Built on top of `openpyxl`.

[![PyPI](https://img.shields.io/pypi/v/excel-schema-engine)](https://pypi.org/project/excel-schema-engine/)
![License](https://img.shields.io/badge/license-MIT-blue)

---

## Installation

```bash
pip install excel-schema-engine
```

or, with Poetry:

```bash
poetry add excel-schema-engine
```

---

## Quickstart

For a complete runnable example, see:

- `examples/quickstart.py` – define a schema, build a workbook, write rows, autosize columns.

---

## Core concepts

- **ExcelSchema**: top‑level object that describes the whole workbook (sheets, styles, author for comments).
- **SheetSchema**: describes a single sheet (its name and list of columns).
- **Column**: describes a logical column with:
  - **key**: dictionary key used when reading/writing rows
  - **header**: text shown in the Excel header
  - **style**: name of a style defined in `ExcelSchema.styles`
  - **comment**: optional `Comment` shown in the header cell
  - **subcolumns**: nested columns for multi‑level headers
- **CellStyle**: strongly‑typed wrapper around `openpyxl` styles (`Font`, `PatternFill`, `Alignment`, `Border`).
- **ExcelErrorsSchema**: defines styles used to mark and highlight errors and fixed values.

All these classes are simple `@dataclass` models and can be constructed in plain Python.

---

## Building Excel templates

See:

- `examples/quickstart.py`

- **Headers**:
  - Single‑row headers are created when a `Column` has no `subcolumns`.
  - Multi‑row headers are created when a `Column` has `subcolumns` (the parent header is merged across all subcolumns).
- **Styles and comments**:
  - Header styles are looked up by name in `schema.styles`.
  - Header comments are created using `Comment(comment, author, width, height)` where `author` comes from `ExcelSchema.author`.

---

## Writing data rows

Use `ExcelBuilder.write_rows` to write dictionaries into the sheet based on the schema.

See:

- `examples/quickstart.py`

- **Missing sheet**: if the given `sheet_name` is not present in the schema, a localized error message is printed and nothing is written.
- **Column mapping**: keys from each `row` dict must match `Column.key` values (including nested subcolumns).

---

## Validating uploaded files

`ExcelValidator` checks that the headers in a user‑submitted file match the schema (order and names).

See:

- `examples/validation_and_reading.py`

`validate_headers` returns a list of human‑readable error messages (already localized).

The validator checks:

- **Number of columns** (expected vs found)
- **Missing columns**
- **Wrong header text** (position, expected text, actual text)

---

## Reading rows as structured data

`ExcelReader` converts worksheet rows into dictionaries using the same `SheetSchema`.

See:

- `examples/validation_and_reading.py`

Each `row` is a `dict` where keys are `Column.key` values and values are cell contents.

---

## Highlighting and marking errors in Excel

Use `ExcelErrorsSchema` and `ExcelErrors` to visually mark problematic cells and rows.

See:

- `examples/mark_errors.py`

- **mark_error**: colors the cell with `error_schema.error_fill` (or a custom `CellStyle`) and optionally adds a localized comment.
- **highlight_row**: colors all cells in a row with `error_schema.highlight_fill` (or custom style).
- **mark_fixed**: re‑styles previously marked cells using `error_schema.fixed_fill` (or custom style) and optionally replaces the comment text.

---

## Localization

The library supports multiple languages for validator and builder messages via the `Language` enum and `ValidatorErrComment`:

- **Supported out of the box**:
  - `Language.EN` – English
  - `Language.RU` – Russian
  - `Language.PL` – Polish
  - `Language.BOBR` – a playful demo language

You can extend or override messages at runtime:

```python
from excel_schema_engine.global_vars import Language, ValidatorErrComment

ValidatorErrComment.messages[Language.FR] = {
    "missing_column": "Colonne manquante : {column}",
    "wrong_header": "Colonne {index} : attendu '{expected}', trouvé '{found}'",
    "columns_count": "Colonnes attendues : {expected}, trouvées : {found}",
    "miss_sheet": "La feuille nommée {sheet_name} est introuvable.",
    "error_prefix": "Erreur : ",
}
```

Messages are formatted using `str.format`, so you can use any placeholders expected by the library (`{column}`, `{index}`, `{expected}`, `{found}`, `{sheet_name}`, etc.).

---

## API overview

Everything you typically need is re‑exported from the package root:

- **Schema and styles**
  - `ExcelSchema`
  - `SheetSchema`
  - `Column`
  - `Comment`
  - `CellStyle`
  - `ExcelErrorsSchema`
- **Builders and readers**
  - `ExcelBuilder`
  - `ExcelReader`
  - `ExcelValidator`
  - `ExcelErrors`
- **Utilities**
  - `autosize_columns`

For more advanced usage, you can also import:

- `Language` and `ValidatorErrComment` from `excel_schema_engine.global_vars`

This should be enough to build templates, validate uploaded files, and work with structured Excel data in your applications.