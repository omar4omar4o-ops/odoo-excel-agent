param(
    [string]$RepoName = "odoo-excel-agent",
    [ValidateSet("public", "private")]
    [string]$Visibility = "public",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Require-Command {
    param([string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw "$Name is required. Install it first, then rerun this script."
    }
    return $cmd.Source
}

function Test-GhCommand {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & gh @Arguments *> $null
        return $LASTEXITCODE -eq 0
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }
}

Require-Command git | Out-Null
Require-Command gh | Out-Null

& gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is not authenticated. Run: gh auth login"
}

$owner = (& gh api user --jq ".login").Trim()
if (-not $owner) {
    throw "Could not detect authenticated GitHub user."
}

$repoFullName = "$owner/$RepoName"
$version = (Select-String -LiteralPath ".\odoo_excel_agent_support.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"').Matches[0].Groups[1].Value
if (-not $version) {
    throw "Could not read APP_VERSION."
}

$releaseTag = "v$version"
$releaseBaseUrl = "https://github.com/$repoFullName/releases/download/$releaseTag"
$env:ODOO_EXCEL_AGENT_UPDATE_BASE_URL = $releaseBaseUrl

if (-not $SkipBuild) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File ".\build_exe.ps1"
}
else {
    & powershell -NoProfile -ExecutionPolicy Bypass -File ".\package_release.ps1"
}

if (-not (Test-Path -LiteralPath ".git")) {
    git init
    git branch -M main
}

$hasOrigin = $false
$remoteNames = @(git remote)
if ($remoteNames -contains "origin") {
    $hasOrigin = $true
}

if (-not $hasOrigin) {
    $visibilityFlag = if ($Visibility -eq "public") { "--public" } else { "--private" }
    $existing = Test-GhCommand @("repo", "view", $repoFullName)
    if (-not $existing) {
        & gh repo create $repoFullName $visibilityFlag --source . --remote origin --disable-wiki --disable-issues
    }
    else {
        git remote add origin "https://github.com/$repoFullName.git"
    }
}

git add .gitignore README.md *.py *.ps1 *.cmd *.json *.txt
git reset -- release OdooExcelAgent.exe build dist .build-venv excel 2>$null

if (git diff --cached --quiet) {
    Write-Host "No source changes to commit."
}
else {
    git commit -m "Publish Odoo Excel Agent $version"
}

git push -u origin main
git tag -f $releaseTag
git push origin $releaseTag --force

$assetPaths = @(
    ".\release\OdooExcelAgent-Windows.zip",
    ".\release\update-manifest.json",
    ".\OdooExcelAgent.exe"
)

$existingRelease = Test-GhCommand @("release", "view", $releaseTag, "--repo", $repoFullName)

if (-not $existingRelease) {
    & gh release create $releaseTag @assetPaths --repo $repoFullName --title "Odoo Excel Agent $version" --notes "Windows release for Odoo Excel Agent $version." --latest
}
else {
    & gh release upload $releaseTag @assetPaths --repo $repoFullName --clobber
}

Write-Host "Published repository: https://github.com/$repoFullName" -ForegroundColor Green
Write-Host "Published release:    https://github.com/$repoFullName/releases/tag/$releaseTag" -ForegroundColor Green
Write-Host "Update URL:           https://api.github.com/repos/$repoFullName/releases/latest" -ForegroundColor Green
