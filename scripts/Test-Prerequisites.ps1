[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-CommandAvailable {
    param([Parameter(Mandatory = $true)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Write-Status {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][bool]$Available,
        [string]$Detail
    )

    if ($Available) {
        Write-Host "[OK]      $Name $Detail" -ForegroundColor Green
    }
    else {
        Write-Host "[MISSING] $Name $Detail" -ForegroundColor Yellow
    }
}

$hasGit = Test-CommandAvailable -Name "git"
$hasNode = Test-CommandAvailable -Name "node"
$hasNpm = Test-CommandAvailable -Name "npm"
$hasWinget = Test-CommandAvailable -Name "winget"
$hasPwsh = Test-CommandAvailable -Name "pwsh"

$nodeVersion = if ($hasNode) { (& node --version) } else { "" }
$npmVersion = if ($hasNpm) { (& npm --version) } else { "" }
$pwshVersion = if ($hasPwsh) { (& pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()') } else { "" }

Write-Host "Outlook-TS prerequisite check"
Write-Host ""
Write-Status -Name "PowerShell 7" -Available $hasPwsh -Detail $pwshVersion
Write-Status -Name "Node.js" -Available $hasNode -Detail $nodeVersion
Write-Status -Name "npm" -Available $hasNpm -Detail $npmVersion
Write-Status -Name "Git" -Available $hasGit -Detail "(optional if using ZIP download)"
Write-Status -Name "winget" -Available $hasWinget -Detail "(optional convenience installer)"
Write-Host ""

if (-not $hasNode -or -not $hasNpm) {
    Write-Host "Node.js/npm are required to run the TypeScript analyzer." -ForegroundColor Yellow
    Write-Host "Install Node.js LTS from https://nodejs.org/en/download, then close and reopen PowerShell."
}

if (-not $hasGit) {
    Write-Host "Git is optional. Use the ZIP download flow in docs/WINDOWS-SETUP.md if Git is not installed."
}

if (-not $hasPwsh) {
    Write-Host "PowerShell 7 is recommended. Install it from https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows."
}

if ($hasNode -and $hasNpm) {
    Write-Host ""
    Write-Host "Next commands:"
    Write-Host "  npm install"
    Write-Host "  npm run build"
    Write-Host "  npm run diagnose:sample"
}
