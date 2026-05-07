[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$UserId,
    [string]$OutputPath = ".\graph-diagnostics.json",
    [switch]$Connect,
    [switch]$IncludeHiddenFolders,
    [int]$LargestFolderCount = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-CollectorError {
    param([Parameter(Mandatory = $true)][string]$Message)
    $script:errors.Add($Message)
}

function Convert-RecipientList {
    param([object[]]$Recipients)

    if ($null -eq $Recipients) {
        return @()
    }

    return @($Recipients | ForEach-Object {
        if ($_.EmailAddress.Address) {
            $_.EmailAddress.Address
        }
        elseif ($_.EmailAddress.Name) {
            $_.EmailAddress.Name
        }
        else {
            $_.ToString()
        }
    })
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Microsoft Graph PowerShell is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Mail -ErrorAction SilentlyContinue

if ($Connect) {
    Connect-MgGraph -Scopes "User.Read", "MailboxSettings.Read", "Mail.ReadBasic", "Mail.Read" -NoWelcome
}

$errors = New-Object System.Collections.Generic.List[string]
$context = Get-MgContext
$settings = $null
$folders = @()
$rules = @()

try {
    $settings = Get-MgUserMailboxSetting -UserId $UserId -ErrorAction Stop
}
catch {
    Add-CollectorError "Get-MgUserMailboxSetting failed for '$UserId': $($_.Exception.Message)"
}

try {
    $folderArgs = @{
        UserId = $UserId
        All = $true
        ErrorAction = "Stop"
    }
    if ($IncludeHiddenFolders) {
        $folderArgs.Add("IncludeHiddenFolders", $true)
    }

    $folders = @(Get-MgUserMailFolder @folderArgs)
}
catch {
    Add-CollectorError "Get-MgUserMailFolder failed for '$UserId': $($_.Exception.Message)"
}

try {
    $rules = @(Get-MgUserMailFolderMessageRule -UserId $UserId -MailFolderId "Inbox" -All -ErrorAction Stop)
}
catch {
    Add-CollectorError "Get-MgUserMailFolderMessageRule failed for '$UserId': $($_.Exception.Message)"
}

$largestFolders = @($folders |
    Sort-Object -Property TotalItemCount -Descending |
    Select-Object -First $LargestFolderCount |
    ForEach-Object {
        [ordered]@{
            displayName = $_.DisplayName
            totalItemCount = $_.TotalItemCount
            unreadItemCount = $_.UnreadItemCount
            childFolderCount = $_.ChildFolderCount
            isHidden = $_.IsHidden
        }
    })

$ruleSummaries = @($rules | ForEach-Object {
    $actions = $_.Actions
    $forwardTo = if ($actions) { Convert-RecipientList -Recipients $actions.ForwardTo } else { @() }
    $redirectTo = if ($actions) { Convert-RecipientList -Recipients $actions.RedirectTo } else { @() }
    $hasForwarding = ($forwardTo.Count -gt 0) -or ($redirectTo.Count -gt 0)
    $hasDeleteOrMove = [bool]($actions -and ($actions.Delete -or $actions.MoveToFolder))

    [ordered]@{
        displayName = $_.DisplayName
        isEnabled = $_.IsEnabled
        sequence = $_.Sequence
        hasForwardingAction = $hasForwarding
        hasDeleteOrMoveAction = $hasDeleteOrMove
        forwardTo = @($forwardTo)
        redirectTo = @($redirectTo)
        moveToFolder = if ($actions) { $actions.MoveToFolder } else { $null }
    }
})

$workingHoursTimeZone = $null
if ($settings -and $settings.WorkingHours -and $settings.WorkingHours.TimeZone) {
    $workingHoursTimeZone = $settings.WorkingHours.TimeZone.Name
}

$automaticRepliesStatus = $null
if ($settings -and $settings.AutomaticRepliesSetting -and $settings.AutomaticRepliesSetting.Status) {
    $automaticRepliesStatus = $settings.AutomaticRepliesSetting.Status.ToString()
}

$diagnostics = [ordered]@{
    collectedAt = (Get-Date).ToUniversalTime().ToString("o")
    targetUserPrincipalName = $UserId
    graph = [ordered]@{
        collectedAt = (Get-Date).ToUniversalTime().ToString("o")
        collectorVersion = "0.1.0"
        errors = @($errors)
        signedInUserPrincipalName = $context.Account
        targetUserPrincipalName = $UserId
        mailboxSettingsAvailable = [bool]$settings
        mailboxSettings = [ordered]@{
            timeZone = if ($settings) { $settings.TimeZone } else { $null }
            dateFormat = if ($settings) { $settings.DateFormat } else { $null }
            timeFormat = if ($settings) { $settings.TimeFormat } else { $null }
            workingHoursTimeZone = $workingHoursTimeZone
            automaticRepliesStatus = $automaticRepliesStatus
        }
        folderCount = @($folders).Count
        hiddenFolderCount = @($folders | Where-Object { $_.IsHidden }).Count
        largestFolders = @($largestFolders)
        inboxRuleCount = @($rules).Count
        enabledInboxRuleCount = @($rules | Where-Object { $_.IsEnabled }).Count
        forwardingRuleCount = @($ruleSummaries | Where-Object { $_.hasForwardingAction }).Count
        suspiciousRules = @($ruleSummaries | Where-Object { $_.hasForwardingAction -or $_.hasDeleteOrMoveAction })
    }
}

$json = $diagnostics | ConvertTo-Json -Depth 12
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$json | Out-File -FilePath $resolvedOutputPath -Encoding utf8
Write-Host "Wrote Graph diagnostics to $resolvedOutputPath"
