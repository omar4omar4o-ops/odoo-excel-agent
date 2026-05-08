Odoo Excel Agent - Windows install

Recommended distribution:
1. Copy the whole release bundle to the target PC, not only OdooExcelAgent.exe.
2. On the target PC, double-click:
   Install Odoo Excel Agent.cmd
3. The installer copies the application to:
   %LOCALAPPDATA%\OdooExcelAgent
4. The setup UI opens automatically.
5. Enter:
   - Odoo URL
   - Odoo database
   - Odoo login
   - Odoo API key
6. Choose the workbook paths:
   - ACHATS LOCAL
   - ACHATS ETRANGER
   - Seller / Previous
7. Click Save.
8. Click Install / Update Agent.

Important:
- On each new Windows user/PC, enter and save the Odoo API key once.
  The key is stored in that user's Windows Credential Manager and is not copied from another PC.
- If status says "Setup required" or "No Odoo API key is available", open the UI, enter the API key, then click Save and Install / Update Agent.
- Odoo database must be the database name only, not a URL.
  Example:
    URL: https://sphe.cloudoo.ma
    Database: sphe.cloudoo.ma
- For normal .xlsx/.xlsm processing, the agent runs in Silent mode and does not open Excel.
- Microsoft Excel desktop is required only for legacy .xls files or Advanced Live mode.
- Internet access to the Odoo server is required.
- If Windows SmartScreen appears, use "More info" then "Run anyway".

Recommended files to ship together:
- OdooExcelAgent.exe
- install_odoo_excel_agent.ps1
- Install Odoo Excel Agent.cmd
- README_WINDOWS_INSTALL.txt

Free updates:
- Option A: upload OdooExcelAgent-Windows.zip to a GitHub Release and paste this URL in the Update tab:
  https://api.github.com/repos/USER/REPO/releases/latest
- Option B: upload OdooExcelAgent-Windows.zip and update-manifest.json to a GitHub Release, then paste the manifest URL.
- Optional before building: set ODOO_EXCEL_AGENT_UPDATE_BASE_URL to your GitHub Release download URL so update-manifest.json contains final URLs.
  Example: https://github.com/USER/REPO/releases/download/v2026.05.08.5
- In the app, open the Update tab and paste the manifest URL.
- The updater verifies SHA-256 when the manifest contains it, keeps config/credentials, replaces the EXE after closing, then reopens the app.
