param(
    [string]$InstallDir = "$env:LOCALAPPDATA\OdooExcelAgent"
)

$ErrorActionPreference = "Stop"

$installPath = [System.IO.Path]::GetFullPath($InstallDir)
$configPath = Join-Path $installPath "config.json"
$startupShortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup\Odoo Excel Agent.lnk"

$processes = Get-CimInstance Win32_Process -Filter "Name = 'pythonw.exe' OR Name = 'python.exe'"
foreach ($process in $processes) {
    $commandLine = [string]$process.CommandLine
    if ($commandLine -like "*odoo_excel_background.py*") {
        Stop-Process -Id $process.ProcessId -Force
    }
}

if (Test-Path -LiteralPath $configPath) {
    try {
        $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        $target = [string]$config.odoo.credential_target
        if (-not [string]::IsNullOrWhiteSpace($target)) {
            & cmdkey /delete:$target | Out-Null
        }
    }
    catch {
    }
}

if (Test-Path -LiteralPath $startupShortcut) {
    Remove-Item -LiteralPath $startupShortcut -Force
}

if (Test-Path -LiteralPath $installPath) {
    Remove-Item -LiteralPath $installPath -Recurse -Force
}

Write-Host "Odoo Excel Agent removed."
