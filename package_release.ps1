$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$exePath = Join-Path $root "OdooExcelAgent.exe"
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "OdooExcelAgent.exe was not found. Build the executable first."
}

$releaseRoot = Join-Path $root "release"
$bundleDir = Join-Path $releaseRoot "OdooExcelAgent-Windows"
$zipPath = Join-Path $releaseRoot "OdooExcelAgent-Windows.zip"
$manifestPath = Join-Path $releaseRoot "update-manifest.json"

if (Test-Path -LiteralPath $bundleDir) {
    Remove-Item -Recurse -Force $bundleDir
}
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -Force $zipPath
}
if (Test-Path -LiteralPath $manifestPath) {
    Remove-Item -Force $manifestPath
}

New-Item -ItemType Directory -Force -Path $bundleDir | Out-Null

$files = @(
    "OdooExcelAgent.exe",
    "install_odoo_excel_agent.ps1",
    "Install Odoo Excel Agent.cmd",
    "README_WINDOWS_INSTALL.txt"
)

foreach ($name in $files) {
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination (Join-Path $bundleDir $name) -Force
}

Compress-Archive -Path (Join-Path $bundleDir "*") -DestinationPath $zipPath -Force

$pythonForVersion = Join-Path $root ".build-venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonForVersion)) {
    $pythonForVersion = "python"
}
$version = (& $pythonForVersion -c "from odoo_excel_agent_support import APP_VERSION; print(APP_VERSION)").Trim()
$baseUrl = $env:ODOO_EXCEL_AGENT_UPDATE_BASE_URL
if (-not $baseUrl) {
    $baseUrl = "https://github.com/YOUR-USER/YOUR-REPO/releases/download/v$version"
}
$baseUrl = $baseUrl.TrimEnd("/")
$zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$exeHash = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
$manifest = [ordered]@{
    version = $version
    notes = "Odoo Excel Agent $version"
    release_url = $baseUrl
    windows_zip = [ordered]@{
        kind = "zip"
        url = "$baseUrl/OdooExcelAgent-Windows.zip"
        sha256 = $zipHash
    }
    windows_exe = [ordered]@{
        kind = "exe"
        url = "$baseUrl/OdooExcelAgent.exe"
        sha256 = $exeHash
    }
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host "Release bundle created:" -ForegroundColor Green
Write-Host "Folder: $bundleDir"
Write-Host "Zip:    $zipPath"
Write-Host "Update manifest: $manifestPath"
