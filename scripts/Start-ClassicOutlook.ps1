param(
    [ValidateSet("Normal", "Safe", "Profiles")]
    $Mode = "Safe",
    [switch]$ForceClassicToggle,
    [switch]$StopExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ClassicOutlookPath {
    $candidates = @(
        "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $command = Get-Command OUTLOOK.EXE -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
        return $command.Source
    }

    return $null
}

if ($StopExisting) {
    Get-Process -Name "OUTLOOK", "olk" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}

if ($ForceClassicToggle) {
    $preferencesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
    if (-not (Test-Path $preferencesPath)) {
        New-Item -Path $preferencesPath -Force | Out-Null
    }

    New-ItemProperty -Path $preferencesPath -Name "UseNewOutlook" -PropertyType DWord -Value 0 -Force | Out-Null
    Write-Host "Set UseNewOutlook=0 for current user."
}

$outlookPath = Get-ClassicOutlookPath
if (-not $outlookPath) {
    throw "Classic Outlook executable was not found in common Office paths or PATH."
}

$arguments = switch ($Mode) {
    "Safe" { "/safe" }
    "Profiles" { "/profiles" }
    default { "" }
}

Write-Host "Launching: $outlookPath $arguments"
$process = if ($arguments) {
    Start-Process -FilePath $outlookPath -ArgumentList $arguments -PassThru
}
else {
    Start-Process -FilePath $outlookPath -PassThru
}

Start-Sleep -Seconds 8
$running = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "OUTLOOK.EXE is running. Process ID: $($running.Id)"
}
else {
    Write-Host "OUTLOOK.EXE started and then exited or crashed." -ForegroundColor Yellow
    Write-Host "Recent Outlook-related Application events:"
    Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = (Get-Date).AddMinutes(-15) } -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.ProviderName -match "Outlook|Office|Application Error|Windows Error Reporting") -or
            ($_.Message -match "OUTLOOK\.EXE|olk\.exe|Microsoft\.OutlookForWindows")
        } |
        Select-Object -First 8 TimeCreated, ProviderName, Id, LevelDisplayName, Message |
        Format-List
}
