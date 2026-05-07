# Windows Setup

This tool is intended for Windows tech machines. It does not require Linux.

## Recommended Tools

- PowerShell 7
- Node.js LTS, which includes `node` and `npm`
- Git, optional but convenient
- winget, optional but convenient
- Chocolatey, optional installer fallback when `winget` is unavailable

Some remote or locked-down machines do not have `git`, `winget`, or `npm`. Use the ZIP workflow below in that case.

## Install Prerequisites With Chocolatey

Use this when `winget` is not available. Open PowerShell as Administrator first.

Install Chocolatey:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://community.chocolatey.org/install.ps1 | iex"
```

Close and reopen PowerShell as Administrator, then install Node.js LTS and PowerShell 7:

```powershell
choco install nodejs-lts --version="24.15.0" -y
choco install powershell-core -y
```

Close and reopen PowerShell again so `PATH` refreshes, then verify:

```powershell
node -v
npm -v
pwsh --version
```

The tool currently supports Node.js 20 or newer. Node.js LTS 24.15.0 works.

## No Git / No winget Workflow

Run this in PowerShell:

```powershell
$zip = "$env:TEMP\Outlook-TS-main.zip"
$dest = "$env:USERPROFILE\Outlook-TS"
Invoke-WebRequest -Uri "https://github.com/grimreach/Outlook-TS/archive/refs/heads/main.zip" -OutFile $zip
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Expand-Archive -Path $zip -DestinationPath $env:USERPROFILE -Force
Rename-Item -Path "$env:USERPROFILE\Outlook-TS-main" -NewName "Outlook-TS"
Set-Location $dest
```

Then check prerequisites:

```powershell
.\scripts\Test-Prerequisites.ps1
```

If `node` or `npm` is missing, install Node.js LTS from:

```text
https://nodejs.org/en/download
```

Or install it with Chocolatey from an elevated PowerShell window:

```powershell
choco install nodejs-lts --version="24.15.0" -y
```

Close and reopen PowerShell after installing Node.js, then run:

```powershell
cd "$env:USERPROFILE\Outlook-TS"
npm install
npm run build
npm run diagnose:sample
```

Start the local web app:

```powershell
npm run web
```

Open this in a browser on the tech PC:

```text
http://127.0.0.1:3000
```

## Git Workflow

Use this if Git is installed:

```powershell
git clone https://github.com/grimreach/Outlook-TS.git
cd Outlook-TS
.\scripts\Test-Prerequisites.ps1
npm install
npm run build
```

Start the local web app:

```powershell
npm run web
```

## Local Outlook Diagnostics

Run this in the affected user's Windows profile:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Collect-OutlookDiagnostics.ps1 -OutputPath .\outlook-diagnostics.json
node .\dist\index.js diagnose --input .\outlook-diagnostics.json
```

## Force Classic Outlook Launch

```powershell
.\scripts\Repair-NewOutlookToggle.ps1 -ForceClassicLaunch
```

## Full Microsoft 365 Diagnostics

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser

.\scripts\Collect-OutlookDiagnostics.ps1 -OutputPath .\outlook-diagnostics.json
.\scripts\Collect-ExchangeOnlineDiagnostics.ps1 -Identity user@contoso.com -Connect -IncludeMailboxPermissions
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -IncludeHiddenFolders

.\scripts\Merge-Diagnostics.ps1 -InputPath .\outlook-diagnostics.json, .\exchange-diagnostics.json, .\graph-diagnostics.json -OutputPath .\combined.json
node .\dist\index.js diagnose --input .\combined.json --markdown .\report.md
```

For a dashboard view, run `npm run web` and upload `combined.json`.
