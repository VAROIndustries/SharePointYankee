<#
.SYNOPSIS
    Lists devices attached to this machine, grouped by class and indented,
    so you can copy a device's Name into the Chameleon scripts as the
    $targetDeviceMatch value.

.DESCRIPTION
    By default shows only devices that are currently present (attached).
    The printed "Name" line is the string the Chameleon scripts match
    against (with wildcards, e.g. "*Dell U2720Q*").

.PARAMETER All
    Include devices that are not currently present (previously attached).

.PARAMETER Filter
    Wildcard to narrow the list, e.g. -Filter *monitor* or -Filter *dock*.

.PARAMETER ShowId
    Also print each device's InstanceId (useful for disambiguating
    identical names).

.EXAMPLE
    .\Get-AttachedDevices.ps1
    .\Get-AttachedDevices.ps1 -Filter *monitor* -ShowId
#>
[CmdletBinding()]
param(
    [switch]$All,
    [string]$Filter,
    [switch]$ShowId
)

$params = @{}
if (-not $All) { $params['PresentOnly'] = $true }

$devices = Get-PnpDevice @params | Where-Object { $_.Name }

if ($Filter) {
    $devices = $devices | Where-Object { $_.Name -like $Filter }
}

if (-not $devices) {
    Write-Host "No matching devices found." -ForegroundColor Yellow
    return
}

$scope = if ($All) { "all known" } else { "currently attached" }
Write-Host ""
Write-Host "Devices ($scope)$(if ($Filter) { " matching '$Filter'" })" -ForegroundColor Cyan
Write-Host ("=" * 60)

$devices |
    Sort-Object Class, Name |
    Group-Object Class |
    Sort-Object Name |
    ForEach-Object {
        $className = if ($_.Name) { $_.Name } else { "(no class)" }
        Write-Host ""
        Write-Host $className -ForegroundColor Green

        $_.Group | Sort-Object Name -Unique | ForEach-Object {
            $status = if ($_.Status -eq 'OK') { '' } else { "  [$($_.Status)]" }
            Write-Host ("    {0}{1}" -f $_.Name, $status)
            if ($ShowId) {
                Write-Host ("        InstanceId: {0}" -f $_.InstanceId) -ForegroundColor DarkGray
            }
        }
    }

Write-Host ""
Write-Host ("Copy a device Name above (wrap it in wildcards, e.g. `"*Dell U2720Q*`")") -ForegroundColor Cyan
Write-Host "into `$targetDeviceMatch in the Chameleon scripts."
Write-Host ""
