# ExcelOps

ExcelOps is a desktop workflow builder and runner for repeatable spreadsheet analysis. It supports CSV/Excel input, filters, sorting, calculated columns, pivot tables, chained VLOOKUP steps, previews, and multi-sheet workbook export.

## Calculated-column formulas

Calculated columns accept direct, row-by-row formulas using column names:

```text
IF(A21 > B32, A21 * B3, 0)
IF(AND(Revenue >= 2500000, SellQty = 0), Revenue, 0)
IF(OR(Status = "Open", Priority = "High"), "Review", "Skip")
```

The leading Excel `=` is optional. Formula references are **column headers**, evaluated row by row; `A21` therefore means a column whose header is `A21`, not spreadsheet cell A21. Use backticks around column names containing spaces, for example `` `Sell Qty` ``. Column matching is case-insensitive, so `b3` can match `B3`. `IF` may omit its third argument, in which case unmatched rows are blank. Formula definitions remain part of the saved workflow.

## Run from source

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

On Windows, activate the environment with `.venv\Scripts\activate`.

## Build the desktop app

macOS/Linux:

```bash
./scripts/build_app.sh
```

Windows PowerShell:

```powershell
.\scripts\build_app.ps1
```

See [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) for signing, testing, packaging, and release instructions.
