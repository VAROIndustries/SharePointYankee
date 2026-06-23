<#
.SYNOPSIS
    Reports storage usage across all SharePoint Online site collections and exports
    the results to CSV.

.DESCRIPTION
    Connects to SharePoint Online using PnP PowerShell, retrieves every site
    collection visible to the account, and produces a per-site storage report that
    includes:

        - Site URL and title
        - Storage used (GB and MB)
        - Storage quota (GB)
        - Percentage of quota consumed
        - Last content-modified date

    Sites that have consumed more than the warning threshold (default 80%) are
    highlighted in yellow in the console output.  All results are exported to CSV.

    This script uses PnP PowerShell (Connect-PnPOnline / Get-PnPTenantSite) which
    requires the SharePoint administrator role or an app registration with
    Sites.FullControl.All.  If you prefer the SharePoint Online Management Shell,
    replace Get-PnPTenantSite with Get-SPOSite -Limit All.

.PARAMETER TenantUrl
    The root SharePoint admin URL for the tenant.
    Example: https://contoso-admin.sharepoint.com

.PARAMETER OutputPath
    Full path for the output CSV file.  Defaults to
    "SPOStorageReport_<timestamp>.csv" in the current directory.

.PARAMETER WarningThreshold
    Percentage of quota used above which a site is flagged in the console output.
    Default: 80.  Valid range: 1-100.

.PARAMETER IncludePersonalSites
    When specified, OneDrive for Business personal site collections (/personal/*)
    are included in the report.  By default they are excluded because tenants with
    many users generate very large row counts.

.EXAMPLE
    .\Get-SPOStorageReport.ps1 -TenantUrl https://contoso-admin.sharepoint.com

    Connects interactively, reports all non-personal site collections, and writes
    the CSV to the current directory.

.EXAMPLE
    .\Get-SPOStorageReport.ps1 -TenantUrl https://contoso-admin.sharepoint.com -OutputPath "C:\Reports\storage.csv" -WarningThreshold 90

    Uses a 90% warning threshold and writes the report to the specified path.

.EXAMPLE
    .\Get-SPOStorageReport.ps1 -TenantUrl https://contoso-admin.sharepoint.com -IncludePersonalSites

    Includes OneDrive for Business personal sites in the report.

.NOTES
    Author  : Geoff Varosky
    Module  : PnP.PowerShell 2.x+
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+
              SharePoint Administrator role (or equivalent app permissions)
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    Storage units: SharePoint Online returns storage in megabytes.  This script
    converts to GB using 1 GB = 1024 MB.

    Site collections with a quota of 0 (unlimited / tenant-default) will show
    "Unlimited" in the QuotaGB column and N/A for PercentUsed.

    App-only auth example:
        Connect-PnPOnline -Url https://contoso-admin.sharepoint.com `
                          -ClientId <AppId> `
                          -CertificatePath <path.pfx> `
                          -CertificatePassword (ConvertTo-SecureString 'pw' -AsPlainText -Force) `
                          -Tenant contoso.onmicrosoft.com

    Required app permissions (SharePoint):
        Sites.FullControl.All   (Application)
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }

[CmdletBinding()]
param (
    [Parameter(
        Mandatory,
        HelpMessage = 'SharePoint admin center URL (e.g., https://contoso-admin.sharepoint.com).'
    )]
    [ValidateNotNullOrEmpty()]
    [string]$TenantUrl,

    [Parameter(
        HelpMessage = 'Full path for the CSV output file.  Defaults to current directory with timestamp.'
    )]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("SPOStorageReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter(
        HelpMessage = 'Percentage used above which a site is highlighted.  Default: 80.'
    )]
    [ValidateRange(1, 100)]
    [int]$WarningThreshold = 80,

    [Parameter(
        HelpMessage = 'Include OneDrive for Business personal site collections in the report.'
    )]
    [switch]$IncludePersonalSites
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Helpers ---
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    switch ($Level) {
        'SUCCESS' { Write-Host   "[$ts][OK]  $Message" -ForegroundColor Green  }
        'WARN'    { Write-Warning          "[$ts][WARN] $Message"               }
        'ERROR'   { Write-Error            "[$ts][ERR]  $Message" -ErrorAction Continue }
        default   { Write-Verbose          "[$ts][INF] $Message"               }
    }
}

function ConvertTo-GB {
    param ([long]$Megabytes)
    return [Math]::Round($Megabytes / 1024, 2)
}
#endregion

try {
    #region --- Connect ---
    Write-Log "Connecting to: $TenantUrl"
    Connect-PnPOnline -Url $TenantUrl -Interactive -ErrorAction Stop
    Write-Log 'Connection established.' -Level SUCCESS
    #endregion

    #region --- Retrieve all site collections ---
    Write-Log 'Retrieving all site collections (this may take a moment for large tenants)...'

    # Retrieve common properties up-front to avoid per-site round trips.
    $allSites = Get-PnPTenantSite -IncludeOneDriveSites:$IncludePersonalSites `
                                  -Detailed `
                                  -ErrorAction Stop

    if (-not $IncludePersonalSites) {
        # Filter out personal sites even if the switch above did not exclude them
        # (behavior varies across PnP module versions).
        $allSites = $allSites | Where-Object { $_.Url -notlike '*/personal/*' }
    }

    $siteCount = ($allSites | Measure-Object).Count
    Write-Log "Found $siteCount site collection(s) to process."
    #endregion

    #region --- Build report ---
    $report  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $current = 0
    $warned  = 0

    foreach ($site in $allSites) {
        $current++

        Write-Progress -Activity 'Collecting storage data' `
                       -Status   "Site $current of $siteCount : $($site.Url)" `
                       -PercentComplete ([int](($current / [Math]::Max(1, $siteCount)) * 100))

        # StorageUsage is in MB.  StorageQuota 0 = unlimited / tenant default.
        $usedMB    = [long]$site.StorageUsageCurrent
        $quotaMB   = [long]$site.StorageQuota
        $usedGB    = ConvertTo-GB -Megabytes $usedMB

        if ($quotaMB -gt 0) {
            $quotaGB    = ConvertTo-GB -Megabytes $quotaMB
            $pctUsed    = [Math]::Round(($usedMB / $quotaMB) * 100, 1)
            $quotaLabel = "$quotaGB GB"
            $pctLabel   = "$pctUsed %"
        }
        else {
            $quotaGB    = 0
            $pctUsed    = 0
            $quotaLabel = 'Unlimited'
            $pctLabel   = 'N/A'
        }

        $overThreshold = ($quotaMB -gt 0) -and ($pctUsed -ge $WarningThreshold)
        if ($overThreshold) { $warned++ }

        $row = [PSCustomObject]@{
            Title            = $site.Title
            Url              = $site.Url
            Template         = $site.Template
            StorageUsedGB    = $usedGB
            StorageUsedMB    = $usedMB
            StorageQuotaGB   = if ($quotaMB -gt 0) { $quotaGB } else { 'Unlimited' }
            PercentUsed      = if ($quotaMB -gt 0) { $pctUsed } else { 'N/A' }
            OverThreshold    = $overThreshold
            LastContentModified = $site.LastContentModifiedDate
            SiteOwner        = $site.Owner
            Status           = $site.Status
        }

        $report.Add($row)

        # Console highlight for over-threshold sites.
        if ($overThreshold) {
            Write-Host ("  [!] {0,-60} {1,6} GB used / {2,6} GB quota ({3})" -f `
                $site.Url, $usedGB, $quotaGB, $pctLabel) -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity 'Collecting storage data' -Completed
    #endregion

    #region --- Export CSV ---
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Log "Created output directory: $outputDir"
    }

    $report | Sort-Object -Property PercentUsed -Descending |
              Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Log "Report exported to: $OutputPath" -Level SUCCESS
    #endregion

    #region --- Console summary ---
    $totalUsedGB  = ($report | Measure-Object -Property StorageUsedGB -Sum).Sum
    $totalUsedGB  = [Math]::Round($totalUsedGB, 2)

    Write-Host ''
    Write-Host '---- SharePoint Online Storage Summary ----' -ForegroundColor Cyan
    Write-Host ("Total sites reported  : {0}"   -f $siteCount)  -ForegroundColor White
    Write-Host ("Sites over {0}% quota : {1}" -f $WarningThreshold, $warned) `
               -ForegroundColor $(if ($warned -gt 0) { 'Yellow' } else { 'White' })
    Write-Host ("Total storage used    : {0} GB" -f $totalUsedGB) -ForegroundColor White
    Write-Host ''
    Write-Host "Output file           : $OutputPath" -ForegroundColor Green
    #endregion
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    Write-Progress -Activity 'Collecting storage data' -Completed -ErrorAction SilentlyContinue
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
