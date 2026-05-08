@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "INSTALLER=%SCRIPT_DIR%install_odoo_excel_agent.ps1"

if not exist "%INSTALLER%" (
    echo install_odoo_excel_agent.ps1 was not found next to this launcher.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER%"
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
    echo.
    echo Installation failed with exit code %EXITCODE%.
    pause
)
exit /b %EXITCODE%
