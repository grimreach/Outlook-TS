[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$ForceClassicLaunch,
    [switch]$StopOutlookProcesses
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $ForceClassicLaunch) {
    Write-Host "No repair was requested. Use -ForceClassicLaunch to set UseNewOutlook to 0 for the current Windows user."
    return
}

if ($StopOutlookProcesses) {
    Get-Process -Name "OUTLOOK", "olk" -ErrorAction SilentlyContinue | Stop-Process -Force
}

$preferencesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
if (-not (Test-Path $preferencesPath)) {
    New-Item -Path $preferencesPath -Force | Out-Null
}

if ($PSCmdlet.ShouldProcess("$preferencesPath\UseNewOutlook", "Set DWORD value to 0")) {
    New-ItemProperty -Path $preferencesPath -Name "UseNewOutlook" -PropertyType DWord -Value 0 -Force | Out-Null
    Write-Host "UseNewOutlook is now set to 0 for the current Windows user."
    Write-Host "Launch classic Outlook from Start Menu as 'Outlook (classic)' or run OUTLOOK.EXE directly."
}
