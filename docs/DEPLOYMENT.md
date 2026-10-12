# ExcelOps Desktop Build and Deployment

## 1. Release model

Build ExcelOps separately on every operating system you support:

- Build the macOS `.app` on macOS.
- Build the Windows application on Windows.
- Build the Linux application on the target Linux distribution.

PyInstaller is not a cross-compiler. A macOS build cannot be used as a Windows build.

## 2. Prepare a release branch

Start from a clean, tested branch:

```bash
git status
git pull --rebase origin main
git checkout -b release/v1.0.0
```

Update the version in `ExcelOps.spec` before a release. Commit and tag the final result:

```bash
git add .
git commit -m "Prepare ExcelOps v1.0.0"
git tag -a v1.0.0 -m "ExcelOps v1.0.0"
```

## 3. Test from source

### macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

### Windows PowerShell

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

Before building, manually verify both home-screen paths:

1. Create Workflow: load CSV and Excel files, configure each feature, preview, save, and export.
2. Run Workflow: select a workflow and its files, finish processing, switch preview sheets, and export.
3. Close and reopen ExcelOps and confirm saved workflows are still available.

## 4. Build the app

### macOS

```bash
chmod +x scripts/build_app.sh
./scripts/build_app.sh
open dist
```

The macOS output is `dist/ExcelOps.app`.

### Windows

Run PowerShell from the repository root:

```powershell
.\scripts\build_app.ps1
```

The Windows output is under `dist\ExcelOps`.

The default build is one-folder rather than one-file. One-folder builds start faster, are easier to diagnose, and are less likely to trigger antivirus false positives.

## 5. Use the final company icon

The in-app BDS/ExcelOps mark is rendered directly by Tkinter and therefore works without an image dependency. The editable vector artwork is in `assets/excelops_logo.svg`.

For the operating-system app icon:

- Export `assets/excelops_logo.svg` as `assets/excelops.icns` for macOS.
- Export it as `assets/excelops.ico` for Windows.
- Re-run the build. `ExcelOps.spec` automatically uses the platform icon when it exists.

## 6. Workflow storage after packaging

Source runs store workflows in the repository's `presets` directory. Packaged builds store user-created workflows in a writable user-data directory:

- macOS: `~/Library/Application Support/ExcelOps/workflows`
- Windows: `%APPDATA%\ExcelOps\workflows`
- Linux: `${XDG_DATA_HOME:-~/.local/share}/ExcelOps/workflows`

Set `EXCELOPS_PRESET_DIR` to override this location for a managed/shared installation.

## 7. macOS signing and notarization

For internal testing, zip the unsigned app:

```bash
ditto -c -k --sequesterRsrc --keepParent dist/ExcelOps.app ExcelOps-macOS-v1.0.0.zip
```

For external distribution, use an Apple Developer ID certificate:

```bash
codesign --deep --force --options runtime \
  --sign "Developer ID Application: YOUR COMPANY (TEAMID)" \
  dist/ExcelOps.app

ditto -c -k --sequesterRsrc --keepParent \
  dist/ExcelOps.app ExcelOps-macOS-v1.0.0.zip

xcrun notarytool submit ExcelOps-macOS-v1.0.0.zip \
  --apple-id "YOUR_APPLE_ID" \
  --team-id "TEAMID" \
  --password "APP_SPECIFIC_PASSWORD" \
  --wait

xcrun stapler staple dist/ExcelOps.app
```

Do not put certificates, passwords, or tokens in the repository.

## 8. Windows signing and distribution

Sign the executable with a trusted code-signing certificate before external distribution. With the Windows SDK installed:

```powershell
signtool sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 `
  /f C:\secure\company-certificate.pfx dist\ExcelOps\ExcelOps.exe
```

Zip the complete `dist\ExcelOps` folder and distribute the archive. Do not send only `ExcelOps.exe`, because one-folder builds depend on the adjacent runtime files.

## 9. Release acceptance checklist

- Launch on a clean machine that does not have Python installed.
- Confirm the company logo, home screen, editor, and workflow-only preview render correctly.
- Load `.csv`, `.xlsx`, and `.xls` samples.
- Run a workflow containing pivot, multiple VLOOKUPs, calculations, and column editing.
- Confirm raw input files remain unchanged.
- Export and open every output sheet in Microsoft Excel.
- Save a workflow, restart the app, and run it again.
- Scan the packaged archive and record its SHA-256 checksum.

Generate a checksum:

```bash
shasum -a 256 ExcelOps-macOS-v1.0.0.zip
```

On Windows:

```powershell
Get-FileHash ExcelOps-Windows-v1.0.0.zip -Algorithm SHA256
```

## 10. Fixing bugs after deployment

Never edit files inside `dist` or inside `ExcelOps.app`. Fix the source, add a regression test or reproducible fixture, rebuild, and ship a higher patch version.

```bash
git checkout -b fix/short-description
# Edit and test the source.
git add .
git commit -m "Fix short description"
./scripts/build_app.sh
```

Release bug fixes as `v1.0.1`, `v1.0.2`, and so on. Keep old release archives so users can roll back if needed.
