$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvDir = Join-Path $root ".build-venv"
$venvPython = Join-Path $venvDir "Scripts\\python.exe"
$venvPyInstaller = Join-Path $venvDir "Scripts\\pyinstaller.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Creating isolated build virtual environment..." -ForegroundColor Cyan
    python -m venv $venvDir
}

Write-Host "Installing/upgrading isolated build dependencies..." -ForegroundColor Cyan
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install --upgrade pyinstaller pywin32 watchdog customtkinter openpyxl

Write-Host "Building OdooExcelAgent.exe with PyInstaller from the isolated environment..." -ForegroundColor Cyan

if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "OdooExcelAgent.exe") { Remove-Item -Force "OdooExcelAgent.exe" }

& $venvPyInstaller --noconfirm --clean .\OdooExcelAgent.spec

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

Write-Host "Build successful. Single-file executable created in the 'dist' folder." -ForegroundColor Green
Move-Item -Path "dist\OdooExcelAgent.exe" -Destination ".\OdooExcelAgent.exe" -Force
Write-Host "Moved OdooExcelAgent.exe to the current folder." -ForegroundColor Green

if (Test-Path ".\package_release.ps1") {
    & powershell -NoProfile -ExecutionPolicy Bypass -File ".\package_release.ps1"
}
