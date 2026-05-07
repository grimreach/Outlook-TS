[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string[]]$InputPath,
    [string]$OutputPath = ".\outlook-diagnostics.merged.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Merge-Object {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Target,
        [Parameter(Mandatory = $true)][pscustomobject]$Source
    )

    foreach ($property in $Source.PSObject.Properties) {
        if ($null -eq $property.Value) {
            continue
        }

        $targetProperty = $Target.PSObject.Properties[$property.Name]
        if ($null -ne $targetProperty -and
            $targetProperty.Value -is [pscustomobject] -and
            $property.Value -is [pscustomobject]) {
            Merge-Object -Target $targetProperty.Value -Source $property.Value
        }
        elseif ($null -ne $targetProperty) {
            $targetProperty.Value = $property.Value
        }
        else {
            $Target | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
        }
    }
}

$merged = [pscustomobject]@{
    schemaVersion = 1
    collectedAt = (Get-Date).ToUniversalTime().ToString("o")
}

foreach ($path in @($InputPath | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
    $item = Get-Content -Path $resolved -Raw | ConvertFrom-Json
    Merge-Object -Target $merged -Source $item
}

$json = $merged | ConvertTo-Json -Depth 12
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$json | Out-File -FilePath $resolvedOutputPath -Encoding utf8
Write-Host "Wrote merged diagnostics to $resolvedOutputPath"
