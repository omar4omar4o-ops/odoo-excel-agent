param(
    [string]$InstallDir = "$env:LOCALAPPDATA\OdooExcelAgent"
)

$ErrorActionPreference = "Stop"

function Get-PythonwPath {
    $pythonw = Get-Command pythonw -ErrorAction SilentlyContinue
    if ($pythonw) {
        return $pythonw.Source
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
        return ""
    }

    $candidate = Join-Path (Split-Path $python.Source -Parent) "pythonw.exe"
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return $python.Source
}

function New-Shortcut {
    param(
        [string]$Path,
        [string]$TargetPath,
        [string]$Arguments,
        [string]$WorkingDirectory
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($Path)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = "$TargetPath,0"
    $shortcut.Save()
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$installRoot = [System.IO.Path]::GetFullPath($InstallDir)
$configPath = Join-Path $installRoot "config.json"
$exeSource = Join-Path $scriptDir "OdooExcelAgent.exe"
$exeTarget = Join-Path $installRoot "OdooExcelAgent.exe"

New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $installRoot "backups") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $installRoot "reports") | Out-Null

if (Test-Path -LiteralPath $exeSource) {
    Copy-Item -LiteralPath $exeSource -Destination $exeTarget -Force

    $desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Odoo Excel Agent.lnk"
    $programsShortcut = Join-Path ([Environment]::GetFolderPath("Programs")) "Odoo Excel Agent.lnk"
    New-Shortcut -Path $desktopShortcut -TargetPath $exeTarget -Arguments "--config `"$configPath`"" -WorkingDirectory $installRoot
    New-Shortcut -Path $programsShortcut -TargetPath $exeTarget -Arguments "--config `"$configPath`"" -WorkingDirectory $installRoot

    Start-Process -FilePath $exeTarget -ArgumentList "--config `"$configPath`"" -WorkingDirectory $installRoot
    Write-Host "Installed Odoo Excel Agent to $installRoot" -ForegroundColor Green
    Write-Host "Desktop and Start Menu shortcuts were created." -ForegroundColor Green
    Write-Host "The setup UI has been opened. Enter Odoo settings, choose workbooks, then click Install / Update Agent." -ForegroundColor Yellow
    exit 0
}

$pythonwPath = Get-PythonwPath
if (-not $pythonwPath) {
    throw "OdooExcelAgent.exe was not found next to this installer, and Python is not available for source mode."
}

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

Ensure-PythonPackage -ImportName "pythoncom" -PackageName "pywin32"
Ensure-PythonPackage -ImportName "watchdog" -PackageName "watchdog"
Ensure-PythonPackage -ImportName "customtkinter" -PackageName "customtkinter"
Ensure-PythonPackage -ImportName "openpyxl" -PackageName "openpyxl"

Start-Process `
    -FilePath $pythonwPath `
    -ArgumentList "`"$scriptDir\odoo_excel_agent_ui.py`" --config `"$configPath`"" `
    -WorkingDirectory $scriptDir

Write-Host "Opened Odoo Excel Agent setup UI in source mode."
Write-Host "Target config path: $configPath"
