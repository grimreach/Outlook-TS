param(
    $OutputPath = ".\outlook-diagnostics.json",
    $EventLogHours = 72
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$collectorVersion = "0.1.1"

Write-Host "Outlook-TS local collector $collectorVersion"

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (-not (Test-Path $Path)) {
        return $null
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }

    $property = $item.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Convert-ToStringOrNull {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    return $Value.ToString()
}

function Get-RegistrySubkeyNames {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return @()
    }

    return @(Get-ChildItem -Path $Path -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName })
}

function Get-OutlookAddins {
    $roots = @(
        "HKCU:\Software\Microsoft\Office\Outlook\Addins",
        "HKLM:\Software\Microsoft\Office\Outlook\Addins",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins"
    )

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($key in Get-ChildItem -Path $root -ErrorAction SilentlyContinue) {
            $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            $items.Add([ordered]@{
                hive = ($root -split "\\")[0]
                name = $key.PSChildName
                friendlyName = Get-ObjectPropertyValue -InputObject $props -Name "FriendlyName"
                description = Get-ObjectPropertyValue -InputObject $props -Name "Description"
                loadBehavior = Get-ObjectPropertyValue -InputObject $props -Name "LoadBehavior"
                manifest = Get-ObjectPropertyValue -InputObject $props -Name "Manifest"
            })
        }
    }

    return @($items)
}

function Get-OutlookEvents {
    param([int]$Hours)

    $start = (Get-Date).AddHours(-1 * $Hours)
    try {
        return @(Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = $start } -ErrorAction Stop |
            Where-Object {
                ($_.ProviderName -match "Outlook|Office|Application Error|Windows Error Reporting") -or
                ($_.Message -match "OUTLOOK\.EXE|olk\.exe|Microsoft\.OutlookForWindows")
            } |
            Select-Object -First 50 |
            ForEach-Object {
                [ordered]@{
                    timeCreated = $_.TimeCreated.ToString("o")
                    providerName = $_.ProviderName
                    id = $_.Id
                    levelDisplayName = $_.LevelDisplayName
                    message = $_.Message
                }
            })
    }
    catch {
        return @([ordered]@{
            timeCreated = (Get-Date).ToString("o")
            providerName = "Outlook-TS Collector"
            id = 0
            levelDisplayName = "Warning"
            message = "Unable to read Application event log: $($_.Exception.Message)"
        })
    }
}

$preferencesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
$profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$outlookRootPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"
$clickToRunPath = "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration"

$newOutlookPackage = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue | Select-Object -First 1
$officeProps = if (Test-Path $clickToRunPath) { Get-ItemProperty -Path $clickToRunPath -ErrorAction SilentlyContinue } else { $null }
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$profiles = Get-RegistrySubkeyNames -Path $profilesPath
$eventLogHoursValue = [int]$EventLogHours

$diagnostics = [ordered]@{
    schemaVersion = 1
    collectedAt = (Get-Date).ToUniversalTime().ToString("o")
    collectorVersion = $collectorVersion
    computerName = $env:COMPUTERNAME
    userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    os = [ordered]@{
        caption = $os.Caption
        version = $os.Version
        buildNumber = $os.BuildNumber
    }
    office = [ordered]@{
        platform = Get-ObjectPropertyValue -InputObject $officeProps -Name "Platform"
        clientVersionToReport = Get-ObjectPropertyValue -InputObject $officeProps -Name "ClientVersionToReport"
        updateChannel = Get-ObjectPropertyValue -InputObject $officeProps -Name "UpdateChannel"
        productReleaseIds = Get-ObjectPropertyValue -InputObject $officeProps -Name "ProductReleaseIds"
    }
    newOutlook = [ordered]@{
        installed = [bool]$newOutlookPackage
        packageFullName = Get-ObjectPropertyValue -InputObject $newOutlookPackage -Name "PackageFullName"
        version = Convert-ToStringOrNull -Value (Get-ObjectPropertyValue -InputObject $newOutlookPackage -Name "Version")
        installLocation = Get-ObjectPropertyValue -InputObject $newOutlookPackage -Name "InstallLocation"
        useNewOutlook = Get-RegistryValue -Path $preferencesPath -Name "UseNewOutlook"
        useNewOutlookRegistryPath = "$preferencesPath\UseNewOutlook"
    }
    classicOutlook = [ordered]@{
        defaultProfile = Get-RegistryValue -Path $outlookRootPath -Name "DefaultProfile"
        profiles = @($profiles)
        profileCount = @($profiles).Count
        executablePaths = @(Get-Command OUTLOOK.EXE -ErrorAction SilentlyContinue | ForEach-Object { $_.Source })
    }
    addins = @(Get-OutlookAddins)
    eventLog = @(Get-OutlookEvents -Hours $eventLogHoursValue)
}

$json = $diagnostics | ConvertTo-Json -Depth 8
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$json | Out-File -FilePath $resolvedOutputPath -Encoding utf8
Write-Host "Wrote Outlook diagnostics to $resolvedOutputPath"
