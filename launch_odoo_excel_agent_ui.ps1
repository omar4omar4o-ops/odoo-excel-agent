$ErrorActionPreference = "Stop"

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

$pythonw = Get-Command pythonw -ErrorAction SilentlyContinue
if ($pythonw) {
    Start-Process -FilePath $pythonw.Source -ArgumentList "`"$PSScriptRoot\odoo_excel_agent_ui.py`"" -WorkingDirectory $PSScriptRoot
}
else {
    Start-Process -FilePath (Get-Command python -ErrorAction Stop).Source -ArgumentList "`"$PSScriptRoot\odoo_excel_agent_ui.py`"" -WorkingDirectory $PSScriptRoot
}
