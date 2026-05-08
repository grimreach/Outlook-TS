[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Identity,
    [int]$OlderThanDays = 730,
    [ValidateSet("DeleteAndAllowRecovery", "MoveToArchive", "PermanentlyDelete")]
    [string]$RetentionAction = "DeleteAndAllowRecovery",
    [string]$TagName = "Outlook-TS Calendar Cleanup 2 Years",
    [string]$PolicyName,
    [switch]$Connect,
    [switch]$Assign,
    [switch]$StartManagedFolderAssistant
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($OlderThanDays -lt 1) {
    throw "OlderThanDays must be 1 or greater."
}

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

Import-Module ExchangeOnlineManagement

if ($Connect) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$mailbox = Get-EXOMailbox -Identity $Identity -Properties RetentionPolicy,ArchiveStatus,AutoExpandingArchiveEnabled -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($PolicyName)) {
    $safeIdentity = ($Identity -replace "[^a-zA-Z0-9._-]", "-")
    $PolicyName = "Outlook-TS Calendar Cleanup - $safeIdentity"
}

$existingTag = Get-RetentionPolicyTag -Identity $TagName -ErrorAction SilentlyContinue
if ($null -eq $existingTag) {
    if ($PSCmdlet.ShouldProcess($TagName, "Create Calendar retention tag")) {
        New-RetentionPolicyTag `
            -Name $TagName `
            -Type Calendar `
            -RetentionEnabled $true `
            -AgeLimitForRetention $OlderThanDays `
            -RetentionAction $RetentionAction `
            -Comment "Outlook-TS calendar cleanup tag for items older than $OlderThanDays days." | Out-Null
    }
}
else {
    if ($PSCmdlet.ShouldProcess($TagName, "Update Calendar retention tag")) {
        Set-RetentionPolicyTag `
            -Identity $TagName `
            -RetentionEnabled $true `
            -AgeLimitForRetention $OlderThanDays `
            -RetentionAction $RetentionAction `
            -Comment "Outlook-TS calendar cleanup tag for items older than $OlderThanDays days." | Out-Null
    }
}

$tagLinks = New-Object System.Collections.Generic.List[string]
if ($mailbox.RetentionPolicy) {
    $currentPolicy = Get-RetentionPolicy -Identity $mailbox.RetentionPolicy.ToString() -ErrorAction Stop
    foreach ($link in @($currentPolicy.RetentionPolicyTagLinks)) {
        $tagLinks.Add($link.ToString())
    }
}

if (-not ($tagLinks -contains $TagName)) {
    $tagLinks.Add($TagName)
}

$existingPolicy = Get-RetentionPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
if ($null -eq $existingPolicy) {
    if ($PSCmdlet.ShouldProcess($PolicyName, "Create retention policy with copied existing tag links")) {
        New-RetentionPolicy -Name $PolicyName -RetentionPolicyTagLinks $tagLinks.ToArray() | Out-Null
    }
}
else {
    if ($PSCmdlet.ShouldProcess($PolicyName, "Update retention policy tag links")) {
        Set-RetentionPolicy -Identity $PolicyName -RetentionPolicyTagLinks $tagLinks.ToArray() | Out-Null
    }
}

$result = [ordered]@{
    identity = $Identity
    currentRetentionPolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy.ToString() } else { $null }
    newRetentionPolicy = $PolicyName
    tagName = $TagName
    tagType = "Calendar"
    olderThanDays = $OlderThanDays
    retentionAction = $RetentionAction
    archiveStatus = if ($mailbox.ArchiveStatus) { $mailbox.ArchiveStatus.ToString() } else { $null }
    autoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
    assigned = [bool]$Assign
    managedFolderAssistantStarted = [bool]$StartManagedFolderAssistant
}

if ($Assign) {
    if ($PSCmdlet.ShouldProcess($Identity, "Assign retention policy '$PolicyName'")) {
        Set-Mailbox -Identity $Identity -RetentionPolicy $PolicyName
    }
}
else {
    Write-Host "Policy was prepared but not assigned. Re-run with -Assign to apply it to $Identity."
}

if ($StartManagedFolderAssistant) {
    if ($PSCmdlet.ShouldProcess($Identity, "Start Managed Folder Assistant")) {
        Start-ManagedFolderAssistant -Identity $Identity
    }
}

$result
