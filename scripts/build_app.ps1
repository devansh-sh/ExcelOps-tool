$ErrorActionPreference = "Stop"

$RootDir = Split-Path -Parent $PSScriptRoot
Set-Location $RootDir

$VenvDir = if ($env:VENV_DIR) { $env:VENV_DIR } else { ".venv-build" }
py -3 -m venv $VenvDir
& "$VenvDir\Scripts\python.exe" -m pip install --upgrade pip
& "$VenvDir\Scripts\python.exe" -m pip install -r requirements-build.txt
& "$VenvDir\Scripts\python.exe" -m PyInstaller --clean --noconfirm ExcelOps.spec

Write-Host "Build complete. Check: $RootDir\dist"
