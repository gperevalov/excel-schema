# Excel Schema Engine

A lightweight Python library for generating, validating, and reading Excel files using a declarative schema.

Define your Excel structure once and use it to:

- generate Excel templates
- validate uploaded files
- read rows as structured data
- highlight errors directly in Excel

Built on top of **openpyxl**.

---

## Features

- Declarative Excel schema
- Multi-level headers
- Cell styles and comments
- Excel validation
- Row parsing
- Error highlighting
- Column autosizing
- Localization support

---

### Localization

The validator supports multiple languages for error messages.

You can define your own translations by implementing a custom message mapper in your project.

If you'd like to add a new localization, feel free to open an issue, submit a pull request, or contact the author.

#### Example:

```python
ValidatorErrComment.messages[Language.FR] = {
    "missing_column": "Colonne manquante: {column}"
}
```

---

## Installation

```bash
pip install excel-schema
```
or
```bash
poetry add excel-schema
```

[![PyPI](https://img.shields.io/pypi/v/excel-schema-engine)](https://pypi.org/project/excel-schema-engine/)

![License](https://img.shields.io/badge/license-MIT-blue)