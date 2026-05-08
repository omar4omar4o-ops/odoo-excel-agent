$ErrorActionPreference = "Stop"

param(
    [string]$Workbook,
    [string]$OdooUrl,
    [string]$OdooDb,
    [string]$OdooLogin,
    [string]$OdooApiKey,
    [string]$RecordUrlExample,
    [string]$ReportDir,
    [string]$BackupDir,
    [switch]$Apply,
    [switch]$VisibleExcel,
    [switch]$PromptSecret
)

function Ensure-PythonPackage {
    param(
        [string]$ImportName,
        [string]$PackageName
    )

    & python -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ImportName') else 1)" | Out-Null
    if ($LASTEXITCODE -eq 0) {
        return
    }

    Write-Host "Installing missing Python package: $PackageName" -ForegroundColor Yellow
    & python -m pip install $PackageName
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install Python package: $PackageName"
    }
}

function Get-ConfigValue {
    param(
        [object]$Object,
        [string[]]$PathParts
    )

    $current = $Object
    foreach ($part in $PathParts) {
        if ($null -eq $current) {
            return $null
        }
        $property = $current.PSObject.Properties[$part]
        if ($null -eq $property) {
            return $null
        }
        $current = $property.Value
    }
    return $current
}

function Resolve-RunnerConfig {
    param([string]$ConfigPath)

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return $null
    }

    $raw = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
    if (-not $raw.Trim()) {
        return $null
    }

    return $raw | ConvertFrom-Json
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$installDir = Join-Path $env:LOCALAPPDATA "OdooExcelAgent"
$configPath = Join-Path $installDir "config.json"
$config = Resolve-RunnerConfig -ConfigPath $configPath

Ensure-PythonPackage -ImportName "pythoncom" -PackageName "pywin32"
Ensure-PythonPackage -ImportName "openpyxl" -PackageName "openpyxl"

if (-not $Workbook) {
    $Workbook = [string](Get-ConfigValue $config @("manual", "last_file"))
}
if (-not $Workbook) {
    $Workbook = [string](Get-ConfigValue $config @("background", "watch_file"))
}
if (-not $Workbook) {
    $Workbook = Join-Path $HOME "Downloads\L'ETAT DES COMMANDES.xlsx"
}

if (-not $OdooUrl) {
    $OdooUrl = [string](Get-ConfigValue $config @("odoo", "url"))
}
if (-not $OdooDb) {
    $OdooDb = [string](Get-ConfigValue $config @("odoo", "db"))
}
if (-not $OdooLogin) {
    $OdooLogin = [string](Get-ConfigValue $config @("odoo", "login"))
}
if (-not $RecordUrlExample) {
    $RecordUrlExample = [string](Get-ConfigValue $config @("odoo", "record_url_example"))
}
if (-not $BackupDir) {
    $BackupDir = [string](Get-ConfigValue $config @("paths", "backup_dir"))
}
if (-not $ReportDir) {
    $ReportDir = [string](Get-ConfigValue $config @("paths", "report_dir"))
}

if (-not $OdooApiKey -and $env:ODOO_API_KEY) {
    $OdooApiKey = $env:ODOO_API_KEY
}

$arguments = @(
    (Join-Path $scriptDir "link_odoo_vendor_bills.py"),
    "--workbook", $Workbook
)

if ($OdooUrl) {
    $arguments += @("--odoo-url", $OdooUrl)
}
if ($OdooDb) {
    $arguments += @("--odoo-db", $OdooDb)
}
if ($OdooLogin) {
    $arguments += @("--odoo-login", $OdooLogin)
}
if ($OdooApiKey) {
    $arguments += @("--odoo-api-key", $OdooApiKey)
}
if ($RecordUrlExample) {
    $arguments += @("--record-url-example", $RecordUrlExample)
}
if ($ReportDir) {
    $arguments += @("--report-dir", $ReportDir)
}
if ($BackupDir) {
    $arguments += @("--backup-dir", $BackupDir)
}
if ($Apply) {
    $arguments += "--apply"
}
if ($VisibleExcel) {
    $arguments += "--visible-excel"
}
if ($VisibleExcel) {
    $arguments += @("--performance-mode", "live")
}
if ($PromptSecret) {
    $arguments += "--prompt-secret"
}

Write-Host "Running Excel-to-Odoo linking..." -ForegroundColor Cyan
Write-Host "Workbook: $Workbook"
if ($Apply) {
    Write-Host "Mode: apply changes" -ForegroundColor Yellow
}
else {
    Write-Host "Mode: dry run" -ForegroundColor Yellow
}

& python @arguments
exit $LASTEXITCODE
