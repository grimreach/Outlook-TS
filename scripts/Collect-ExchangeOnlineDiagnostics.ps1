[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Identity,
    [string]$OutputPath = ".\exchange-diagnostics.json",
    [switch]$Connect,
    [switch]$IncludeMailboxPermissions
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-CollectorError {
    param([Parameter(Mandatory = $true)][string]$Message)
    $script:errors.Add($Message)
}

function Convert-ExchangeByteStringToBytes {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = $Value.ToString()
    if ($text -match "\(([\d,]+)\s+bytes\)") {
        return [int64](($Matches[1]) -replace ",", "")
    }

    if ($text -match "([\d.]+)\s*(GB|MB|KB|B)") {
        $number = [double]$Matches[1]
        switch ($Matches[2]) {
            "GB" { return [int64]($number * 1GB) }
            "MB" { return [int64]($number * 1MB) }
            "KB" { return [int64]($number * 1KB) }
            "B" { return [int64]$number }
        }
    }

    return $null
}

function Convert-UnlimitedQuotaToBytes {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value.ToString() -eq "Unlimited") {
        return $null
    }

    if ($Value.PSObject.Properties.Name -contains "Value" -and $null -ne $Value.Value) {
        return Convert-ExchangeByteStringToBytes -Value $Value.Value
    }

    return Convert-ExchangeByteStringToBytes -Value $Value
}

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

Import-Module ExchangeOnlineManagement

if ($Connect) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$errors = New-Object System.Collections.Generic.List[string]
$mailbox = $null
$statistics = $null
$cas = $null
$permissions = @()
$calendarFolders = @()
$calendarFolderSettings = $null
$calendarFolderPermissions = @()

try {
    $mailbox = Get-EXOMailbox -Identity $Identity -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails,AccountDisabled,LitigationHoldEnabled,ArchiveStatus,ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward,HiddenFromAddressListsEnabled,RetentionPolicy,RoleAssignmentPolicy,OWAMailboxPolicy,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota -ErrorAction Stop
}
catch {
    Add-CollectorError "Get-EXOMailbox failed for '$Identity': $($_.Exception.Message)"
}

if ($null -ne $mailbox) {
    try {
        $statistics = Get-EXOMailboxStatistics -Identity $Identity -Properties TotalItemSize,ItemCount,DeletedItemCount,TotalDeletedItemSize,LastLogonTime -ErrorAction Stop
    }
    catch {
        Add-CollectorError "Get-EXOMailboxStatistics failed for '$Identity': $($_.Exception.Message)"
    }

    try {
        $cas = Get-CASMailbox -Identity $Identity -ErrorAction Stop
    }
    catch {
        Add-CollectorError "Get-CASMailbox failed for '$Identity': $($_.Exception.Message)"
    }

    if ($IncludeMailboxPermissions) {
        try {
            $permissions = @(Get-EXOMailboxPermission -Identity $Identity -ErrorAction Stop |
                Where-Object {
                    -not $_.IsInherited -and
                    $_.User -notmatch "NT AUTHORITY\\SELF" -and
                    ($_.AccessRights -contains "FullAccess")
                })
        }
        catch {
            Add-CollectorError "Get-EXOMailboxPermission failed for '$Identity': $($_.Exception.Message)"
        }
    }

    try {
        $calendarFolders = @(Get-EXOMailboxFolderStatistics -Identity $Identity -FolderScope Calendar -IncludeOldestAndNewestItems -ErrorAction Stop)
    }
    catch {
        Add-CollectorError "Get-EXOMailboxFolderStatistics calendar scope failed for '$Identity': $($_.Exception.Message)"
    }

    try {
        $calendarFolderSettings = Get-MailboxCalendarFolder -Identity "${Identity}:\Calendar" -ErrorAction Stop
    }
    catch {
        Add-CollectorError "Get-MailboxCalendarFolder failed for '$Identity`:\Calendar': $($_.Exception.Message)"
    }

    try {
        $calendarFolderPermissions = @(Get-EXOMailboxFolderPermission -Identity "${Identity}:\Calendar" -ErrorAction Stop)
    }
    catch {
        try {
            $calendarFolderPermissions = @(Get-MailboxFolderPermission -Identity "${Identity}:\Calendar" -ErrorAction Stop)
        }
        catch {
            Add-CollectorError "Get mailbox calendar folder permission failed for '$Identity`:\Calendar': $($_.Exception.Message)"
        }
    }
}

$totalItemSize = if ($null -ne $statistics) { $statistics.TotalItemSize } else { $null }
$totalDeletedItemSize = if ($null -ne $statistics) { $statistics.TotalDeletedItemSize } else { $null }
$lastLogonTime = if ($null -ne $statistics) { $statistics.LastLogonTime } else { $null }

$prohibitSendReceiveQuota = if ($null -ne $mailbox) { $mailbox.ProhibitSendReceiveQuota } else { $null }
$prohibitSendQuota = if ($null -ne $mailbox) { $mailbox.ProhibitSendQuota } else { $null }
$issueWarningQuota = if ($null -ne $mailbox) { $mailbox.IssueWarningQuota } else { $null }

$totalBytes = Convert-ExchangeByteStringToBytes -Value $totalItemSize
$quotaBytes = Convert-UnlimitedQuotaToBytes -Value $prohibitSendReceiveQuota
if ($null -eq $quotaBytes) {
    $quotaBytes = Convert-UnlimitedQuotaToBytes -Value $prohibitSendQuota
}

$quotaUsedPercent = $null
if ($null -ne $totalBytes -and $null -ne $quotaBytes -and $quotaBytes -gt 0) {
    $quotaUsedPercent = [math]::Round(($totalBytes / $quotaBytes) * 100, 2)
}

$calendarFolderSummaries = @($calendarFolders | ForEach-Object {
    [ordered]@{
        name = $_.Name
        folderPath = if ($_.FolderPath) { $_.FolderPath.ToString() } else { $null }
        folderType = if ($_.FolderType) { $_.FolderType.ToString() } else { $null }
        itemsInFolder = $_.ItemsInFolder
        itemsInFolderAndSubfolders = $_.ItemsInFolderAndSubfolders
        folderSize = if ($_.FolderSize) { $_.FolderSize.ToString() } else { $null }
        folderAndSubfolderSize = if ($_.FolderAndSubfolderSize) { $_.FolderAndSubfolderSize.ToString() } else { $null }
        oldestItemReceivedDate = if ($_.OldestItemReceivedDate) { ([datetime]$_.OldestItemReceivedDate).ToUniversalTime().ToString("o") } else { $null }
        newestItemReceivedDate = if ($_.NewestItemReceivedDate) { ([datetime]$_.NewestItemReceivedDate).ToUniversalTime().ToString("o") } else { $null }
    }
})

$defaultCalendarSummary = @($calendarFolderSummaries | Where-Object {
    $_.folderPath -eq "/Calendar" -or $_.name -eq "Calendar"
} | Select-Object -First 1)

$diagnostics = [ordered]@{
    collectedAt = (Get-Date).ToUniversalTime().ToString("o")
    targetUserPrincipalName = $Identity
    exchangeOnline = [ordered]@{
        collectedAt = (Get-Date).ToUniversalTime().ToString("o")
        collectorVersion = "0.1.0"
        errors = @($errors)
        mailboxFound = [bool]$mailbox
        primarySmtpAddress = if ($mailbox) { $mailbox.PrimarySmtpAddress.ToString() } else { $null }
        displayName = if ($mailbox) { $mailbox.DisplayName } else { $null }
        recipientTypeDetails = if ($mailbox) { $mailbox.RecipientTypeDetails.ToString() } else { $null }
        accountDisabled = if ($mailbox) { $mailbox.AccountDisabled } else { $null }
        litigationHoldEnabled = if ($mailbox) { $mailbox.LitigationHoldEnabled } else { $null }
        archiveStatus = if ($mailbox) { $mailbox.ArchiveStatus.ToString() } else { $null }
        forwardingSmtpAddress = if ($mailbox -and $mailbox.ForwardingSmtpAddress) { $mailbox.ForwardingSmtpAddress.ToString() } else { $null }
        forwardingAddress = if ($mailbox -and $mailbox.ForwardingAddress) { $mailbox.ForwardingAddress.ToString() } else { $null }
        deliverToMailboxAndForward = if ($mailbox) { $mailbox.DeliverToMailboxAndForward } else { $null }
        hiddenFromAddressListsEnabled = if ($mailbox) { $mailbox.HiddenFromAddressListsEnabled } else { $null }
        retentionPolicy = if ($mailbox -and $mailbox.RetentionPolicy) { $mailbox.RetentionPolicy.ToString() } else { $null }
        roleAssignmentPolicy = if ($mailbox -and $mailbox.RoleAssignmentPolicy) { $mailbox.RoleAssignmentPolicy.ToString() } else { $null }
        owaMailboxPolicy = if ($mailbox -and $mailbox.OWAMailboxPolicy) { $mailbox.OWAMailboxPolicy.ToString() } else { $null }
        issueWarningQuota = if ($issueWarningQuota) { $issueWarningQuota.ToString() } else { $null }
        prohibitSendQuota = if ($prohibitSendQuota) { $prohibitSendQuota.ToString() } else { $null }
        prohibitSendReceiveQuota = if ($prohibitSendReceiveQuota) { $prohibitSendReceiveQuota.ToString() } else { $null }
        totalItemSize = if ($totalItemSize) { $totalItemSize.ToString() } else { $null }
        totalItemSizeBytes = $totalBytes
        quotaUsedPercent = $quotaUsedPercent
        itemCount = if ($statistics) { $statistics.ItemCount } else { $null }
        deletedItemCount = if ($statistics) { $statistics.DeletedItemCount } else { $null }
        totalDeletedItemSize = if ($totalDeletedItemSize) { $totalDeletedItemSize.ToString() } else { $null }
        lastLogonTime = if ($lastLogonTime) { ([datetime]$lastLogonTime).ToUniversalTime().ToString("o") } else { $null }
        mailboxPermissionSummary = [ordered]@{
            fullAccessDelegates = @($permissions | Select-Object -First 25 | ForEach-Object { $_.User.ToString() })
            nonInheritedPermissionCount = @($permissions).Count
        }
        calendarFolders = @($calendarFolderSummaries)
        defaultCalendar = if (@($defaultCalendarSummary).Count -gt 0) { $defaultCalendarSummary[0] } else { $null }
        calendarFolderSettings = [ordered]@{
            identity = if ($calendarFolderSettings) { $calendarFolderSettings.Identity.ToString() } else { $null }
            publishEnabled = if ($calendarFolderSettings) { $calendarFolderSettings.PublishEnabled } else { $null }
            detailLevel = if ($calendarFolderSettings -and $calendarFolderSettings.DetailLevel) { $calendarFolderSettings.DetailLevel.ToString() } else { $null }
            searchableUrlEnabled = if ($calendarFolderSettings) { $calendarFolderSettings.SearchableUrlEnabled } else { $null }
        }
        calendarFolderPermissions = @($calendarFolderPermissions | ForEach-Object {
            [ordered]@{
                user = if ($_.User) { $_.User.ToString() } else { $null }
                accessRights = @($_.AccessRights | ForEach-Object { $_.ToString() })
                sharingPermissionFlags = @($_.SharingPermissionFlags | ForEach-Object { $_.ToString() })
            }
        })
        cas = [ordered]@{
            owaEnabled = if ($cas) { $cas.OWAEnabled } else { $null }
            mapiEnabled = if ($cas) { $cas.MAPIEnabled } else { $null }
            imapEnabled = if ($cas) { $cas.ImapEnabled } else { $null }
            popEnabled = if ($cas) { $cas.PopEnabled } else { $null }
            activeSyncEnabled = if ($cas) { $cas.ActiveSyncEnabled } else { $null }
            smtpClientAuthenticationDisabled = if ($cas) { $cas.SmtpClientAuthenticationDisabled } else { $null }
            ewsEnabled = if ($cas) { $cas.EwsEnabled } else { $null }
        }
    }
}

$json = $diagnostics | ConvertTo-Json -Depth 8
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$json | Out-File -FilePath $resolvedOutputPath -Encoding utf8
Write-Host "Wrote Exchange Online diagnostics to $resolvedOutputPath"
