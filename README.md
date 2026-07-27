# ExcelOps

ExcelOps is a desktop workflow builder and runner for repeatable spreadsheet analysis. It supports CSV/Excel input, filters, sorting, calculated columns, pivot tables, chained VLOOKUP steps, previews, and multi-sheet workbook export.

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
