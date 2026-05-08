[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$UserId,
    [int]$OlderThanYears = 2,
    [string]$OutputPath = ".\old-calendar-items.json",
    [string]$CsvPath = ".\old-calendar-items.csv",
    [switch]$Connect,
    [switch]$Delete,
    [switch]$IncludeRecurring,
    [int]$PageSize = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($Name)) {
        return $InputObject[$Name]
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

if ($OlderThanYears -lt 1) {
    throw "OlderThanYears must be 1 or greater."
}

if ($PageSize -lt 1 -or $PageSize -gt 999) {
    throw "PageSize must be between 1 and 999."
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Microsoft.Graph.Authentication is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}

Import-Module Microsoft.Graph.Authentication

if ($Connect) {
    Connect-MgGraph -Scopes "Calendars.ReadWrite", "User.Read.All" -NoWelcome
}

$cutoff = (Get-Date).AddYears(-1 * $OlderThanYears).ToUniversalTime()
$start = [datetime]"1900-01-01T00:00:00Z"
$encodedUserId = [uri]::EscapeDataString($UserId)
$select = "id,subject,start,end,type,isOrganizer,organizer,seriesMasterId,lastModifiedDateTime"
$uri = "/v1.0/users/$encodedUserId/calendarView?startDateTime=$($start.ToString("o"))&endDateTime=$($cutoff.ToString("o"))&`$select=$select&`$top=$PageSize&`$orderby=start/dateTime"

$items = New-Object System.Collections.Generic.List[object]

Write-Host "Finding default-calendar items for $UserId older than $($cutoff.ToString("yyyy-MM-dd"))..."
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    foreach ($event in @(Get-ObjectPropertyValue -InputObject $response -Name "value")) {
        $eventType = Get-ObjectPropertyValue -InputObject $event -Name "type"
        if (-not $IncludeRecurring -and $eventType -ne "singleInstance") {
            continue
        }

        $startValue = Get-ObjectPropertyValue -InputObject $event -Name "start"
        $endValue = Get-ObjectPropertyValue -InputObject $event -Name "end"
        $organizerValue = Get-ObjectPropertyValue -InputObject $event -Name "organizer"
        $organizerEmailAddress = Get-ObjectPropertyValue -InputObject $organizerValue -Name "emailAddress"

        $items.Add([ordered]@{
            id = Get-ObjectPropertyValue -InputObject $event -Name "id"
            subject = Get-ObjectPropertyValue -InputObject $event -Name "subject"
            type = $eventType
            start = Get-ObjectPropertyValue -InputObject $startValue -Name "dateTime"
            startTimeZone = Get-ObjectPropertyValue -InputObject $startValue -Name "timeZone"
            end = Get-ObjectPropertyValue -InputObject $endValue -Name "dateTime"
            endTimeZone = Get-ObjectPropertyValue -InputObject $endValue -Name "timeZone"
            isOrganizer = Get-ObjectPropertyValue -InputObject $event -Name "isOrganizer"
            organizer = Get-ObjectPropertyValue -InputObject $organizerEmailAddress -Name "address"
            seriesMasterId = Get-ObjectPropertyValue -InputObject $event -Name "seriesMasterId"
            lastModifiedDateTime = Get-ObjectPropertyValue -InputObject $event -Name "lastModifiedDateTime"
        })
    }

    $uri = Get-ObjectPropertyValue -InputObject $response -Name "@odata.nextLink"
} while ($null -ne $uri)

$itemArray = @($items)
$summary = [ordered]@{
    collectedAt = (Get-Date).ToUniversalTime().ToString("o")
    userId = $UserId
    cutoffUtc = $cutoff.ToString("o")
    includeRecurring = [bool]$IncludeRecurring
    deleteRequested = [bool]$Delete
    itemCount = $itemArray.Count
    items = $itemArray
}

$summary | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
$itemArray | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Found $($itemArray.Count) item(s)."
Write-Host "Wrote JSON preview to $OutputPath"
Write-Host "Wrote CSV preview to $CsvPath"

if (-not $Delete) {
    Write-Host "Preview only. Re-run with -Delete after reviewing the export."
    return
}

$deleted = 0
foreach ($item in $itemArray) {
    $target = "$UserId calendar item '$($item.subject)' starting $($item.start)"
    if ($PSCmdlet.ShouldProcess($target, "Delete calendar item")) {
        $eventId = [uri]::EscapeDataString($item.id)
        Invoke-MgGraphRequest -Method DELETE -Uri "/v1.0/users/$encodedUserId/events/$eventId"
        $deleted++
    }
}

Write-Host "Deleted $deleted item(s)."
