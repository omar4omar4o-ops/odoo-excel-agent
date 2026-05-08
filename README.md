# Odoo Excel Agent

Windows desktop agent that links selected Excel workbooks to Odoo purchase orders.

## Main Features

- Silent background monitoring for selected Excel workbooks.
- ACHATS LOCAL lookup with row fallback: `N°FACTURE` then `N commandes`.
- ACHATS ETRANGER lookup using only `MTT DE FACTURE`.
- Seller / previous workbook keeps the legacy lookup flow.
- Writes hyperlinks to closed `.xlsx/.xlsm` files with `openpyxl`.
- Keeps API keys in Windows Credential Manager.
- Supports free updates from GitHub Releases.

## Windows Install

Use the packaged release bundle:

1. Download `OdooExcelAgent-Windows.zip` from Releases.
2. Extract it.
3. Run `Install Odoo Excel Agent.cmd`.
4. Enter Odoo settings and choose workbook paths.
5. Click `Save`, then `Install / Update Agent`.

## Updates

The app can update itself from a GitHub Release:

- The app now includes the official update URL automatically. If you need to reset it, use:
  `https://api.github.com/repos/OWNER/REPO/releases/latest`
- The release must contain `OdooExcelAgent-Windows.zip`.
- If `update-manifest.json` is also uploaded, the app can verify SHA-256 before installing.

## Build

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Outputs:

- `OdooExcelAgent.exe`
- `release\OdooExcelAgent-Windows.zip`
- `release\update-manifest.json`

## Publish

After installing and authenticating GitHub CLI:

```powershell
gh auth login
powershell -NoProfile -ExecutionPolicy Bypass -File .\publish_github_release.ps1 -RepoName odoo-excel-agent -Visibility public
```

Public visibility is recommended if update downloads should work on other PCs without a GitHub token.
