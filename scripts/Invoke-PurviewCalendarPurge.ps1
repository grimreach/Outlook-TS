[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [string]$FolderName = "Calendar",
    [string]$SearchName,
    [int]$OlderThanYears = 2,
    [datetime]$StartDate = [datetime]"1900-01-01",
    [ValidateSet("SoftDelete", "HardDelete")]
    [string]$PurgeType = "HardDelete",
    [switch]$Connect,
    [switch]$Purge,
    [int]$PollSeconds = 15,
    [int]$MaxPurgePasses = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Convert-MailboxFolderIdToPurviewFolderId {
    param([Parameter(Mandatory = $true)][string]$FolderId)

    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF")
    $folderIdBytes = [Convert]::FromBase64String($FolderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0

    $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
    }

    return $encoding.GetString($indexIdBytes)
}

function Wait-ComplianceSearchComplete {
    param([Parameter(Mandatory = $true)][string]$Identity)

    do {
        Start-Sleep -Seconds $PollSeconds
        $search = Get-ComplianceSearch -Identity $Identity
        Write-Host "Search status: $($search.Status); items: $($search.Items)"
    } while ($search.Status -in @("Starting", "InProgress", "Stopping"))

    return $search
}

function Wait-ComplianceSearchActionComplete {
    param([Parameter(Mandatory = $true)][string]$Identity)

    do {
        Start-Sleep -Seconds $PollSeconds
        $action = Get-ComplianceSearchAction -Identity $Identity
        Write-Host "Purge action status: $($action.Status)"
    } while ($action.Status -in @("Starting", "InProgress", "Stopping"))

    return $action
}

if ($OlderThanYears -lt 1) {
    throw "OlderThanYears must be 1 or greater."
}

if ($PollSeconds -lt 5) {
    throw "PollSeconds must be at least 5."
}

if ($MaxPurgePasses -lt 1) {
    throw "MaxPurgePasses must be 1 or greater."
}

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

Import-Module ExchangeOnlineManagement

if ($Connect) {
    Connect-ExchangeOnline -ShowBanner:$false
    Connect-IPPSSession
}

$cutoff = (Get-Date).AddYears(-1 * $OlderThanYears).Date
if ([string]::IsNullOrWhiteSpace($SearchName)) {
    $safeMailbox = $Mailbox -replace "[^a-zA-Z0-9._-]", "-"
    $SearchName = "OutlookTS-CalendarPurge-$safeMailbox"
}

Write-Host "Fetching folder ID for '$FolderName' in $Mailbox..." -ForegroundColor Cyan
$folderStats = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope Calendar |
    Where-Object { $_.Name -eq $FolderName -or $_.FolderPath -eq "/$FolderName" } |
    Select-Object -First 1

if ($null -eq $folderStats) {
    throw "Could not find calendar folder '$FolderName' for $Mailbox."
}

$purviewFolderId = Convert-MailboxFolderIdToPurviewFolderId -FolderId $folderStats.FolderId
$query = 'kind:meetings AND start>="{0}" AND start<"{1}" AND folderid:{2}' -f $StartDate.ToString("yyyy-MM-dd"), $cutoff.ToString("yyyy-MM-dd"), $purviewFolderId

Write-Host "Target folder path: $($folderStats.FolderPath)" -ForegroundColor Green
Write-Host "Purview folder ID: $purviewFolderId" -ForegroundColor Green
Write-Host "Query: $query" -ForegroundColor Cyan

Write-Host "Cleaning up prior search/action named $SearchName..." -ForegroundColor Yellow
Remove-ComplianceSearchAction -Identity "${SearchName}_Purge" -Confirm:$false -ErrorAction SilentlyContinue
Remove-ComplianceSearch -Identity $SearchName -Confirm:$false -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

if ($PSCmdlet.ShouldProcess($SearchName, "Create targeted Purview compliance search")) {
    New-ComplianceSearch -Name $SearchName -ExchangeLocation $Mailbox -ContentMatchQuery $query | Out-Null
}

Write-Host "Starting compliance search..." -ForegroundColor Cyan
Start-ComplianceSearch -Identity $SearchName
$search = Wait-ComplianceSearchComplete -Identity $SearchName

if ($search.Status -ne "Completed") {
    throw "Compliance search ended with status '$($search.Status)'. Review the search in Purview before purging."
}

$remainingItems = [int]$search.Items
Write-Host "Matched calendar items older than $($cutoff.ToString("yyyy-MM-dd")): $remainingItems" -ForegroundColor Green

if ($remainingItems -eq 0) {
    Write-Host "No matching items found."
    return
}

if (-not $Purge) {
    Write-Host "Preview only. Re-run with -Purge to delete matched items." -ForegroundColor Yellow
    return
}

Write-Host "Purging with PurgeType=$PurgeType. Microsoft Purview purge actions remove a limited batch per mailbox per action, so this loops until the search returns zero items." -ForegroundColor Yellow

$pass = 0
do {
    $pass++
    if ($pass -gt $MaxPurgePasses) {
        throw "Reached MaxPurgePasses=$MaxPurgePasses with $remainingItems item(s) still matching."
    }

    Remove-ComplianceSearchAction -Identity "${SearchName}_Purge" -Confirm:$false -ErrorAction SilentlyContinue

    if ($PSCmdlet.ShouldProcess($SearchName, "Run Purview purge pass $pass using $PurgeType")) {
        Write-Host "Purge pass $pass..." -ForegroundColor Cyan
        New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType $PurgeType -Confirm:$false | Out-Null
        $action = Wait-ComplianceSearchActionComplete -Identity "${SearchName}_Purge"
        if ($action.Status -ne "Completed") {
            throw "Purge action ended with status '$($action.Status)'. Review the action in Purview before continuing."
        }
    }

    Write-Host "Refreshing search count..." -ForegroundColor Cyan
    Start-ComplianceSearch -Identity $SearchName
    $search = Wait-ComplianceSearchComplete -Identity $SearchName
    $remainingItems = [int]$search.Items
    Write-Host "Remaining matching items: $remainingItems" -ForegroundColor Yellow
} while ($remainingItems -gt 0)

Write-Host "Success. All targeted calendar items older than $($cutoff.ToString("yyyy-MM-dd")) were purged from $Mailbox." -ForegroundColor Green
