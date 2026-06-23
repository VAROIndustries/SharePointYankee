<#
.SYNOPSIS
    Bulk-creates SharePoint Communication or Team sites from a CSV input file.

.DESCRIPTION
    Reads a CSV file specifying sites to create and provisions each one via
    PnP PowerShell.  Supports both Communication sites and Microsoft 365 Group-
    backed Team sites.

    Key behaviors:
        - Skips sites that already exist (resume capability — safe to re-run
          after partial failures without duplicating sites)
        - Progress bar shows per-site status across the full batch
        - All errors are captured in an in-memory log and optionally written to
          a separate error CSV at the end of the run
        - Full -WhatIf support — no sites are created when -WhatIf is specified
        - Detailed verbose output at each step

    Required CSV columns (case-insensitive):
        Title       — Display name of the site
        Url         — Relative or full URL of the new site
                      Relative: /sites/Marketing  →  resolved against TenantUrl
                      Full:     https://contoso.sharepoint.com/sites/Marketing
        Template    — CommunicationSite or TeamSite
        Owner       — UPN of the primary site owner (e.g., jsmith@contoso.com)
        Description — (Optional) Site description

    CSV example:
        Title,Url,Template,Owner,Description
        "HR Portal",/sites/HR,CommunicationSite,admin@contoso.com,"Human Resources intranet"
        "Finance Team",/sites/Finance,TeamSite,cfo@contoso.com,"Finance department team site"

.PARAMETER InputFile
    Path to the CSV file containing the site definitions.

.PARAMETER TenantUrl
    The SharePoint Online admin center URL or the root site URL for the tenant.
    Example: https://contoso-admin.sharepoint.com

.PARAMETER ErrorLogPath
    Optional path for a CSV file that captures any sites that failed to create.
    Defaults to "<InputFile basename>_Errors_<timestamp>.csv" alongside the
    input file.

.PARAMETER ThrottleDelaySeconds
    Number of seconds to wait between site creation requests.  SharePoint Online
    throttles bulk provisioning; a small delay reduces the chance of hitting
    rate limits.  Default: 5.

.EXAMPLE
    .\New-BulkSharePointSites.ps1 -InputFile "C:\Sites\new-sites.csv" -TenantUrl https://contoso-admin.sharepoint.com

    Creates all sites defined in the CSV.  Skips any that already exist.

.EXAMPLE
    .\New-BulkSharePointSites.ps1 -InputFile "C:\Sites\new-sites.csv" -TenantUrl https://contoso-admin.sharepoint.com -WhatIf

    Shows what would be created without making any changes.

.EXAMPLE
    .\New-BulkSharePointSites.ps1 -InputFile "C:\Sites\new-sites.csv" -TenantUrl https://contoso-admin.sharepoint.com -ThrottleDelaySeconds 10

    Uses a 10-second delay between each site creation request.

.NOTES
    Author  : Geoff Varosky
    Module  : PnP.PowerShell 2.x+
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+
              SharePoint Administrator role (or equivalent app permissions)
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    Communication sites are created with New-PnPSite -Type CommunicationSite.
    Team sites (Microsoft 365 Group-backed) are created with
    New-PnPSite -Type TeamSite.  Classic team sites (STS#0) are not supported
    by this script because Microsoft recommends Modern sites for all new
    provisioning.

    After creation, site properties such as external sharing level, hub
    association, and permissions should be configured via separate scripts or
    the SharePoint admin center.
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param (
    [Parameter(
        Mandatory,
        HelpMessage = 'Path to the CSV file containing site definitions.'
    )]
    [ValidateScript({
        if (-not (Test-Path -Path $_ -PathType Leaf)) {
            throw "InputFile not found: $_"
        }
        if ([System.IO.Path]::GetExtension($_) -ne '.csv') {
            throw "InputFile must be a .csv file: $_"
        }
        $true
    })]
    [string]$InputFile,

    [Parameter(
        Mandatory,
        HelpMessage = 'SharePoint admin center URL (e.g., https://contoso-admin.sharepoint.com).'
    )]
    [ValidateNotNullOrEmpty()]
    [string]$TenantUrl,

    [Parameter(
        HelpMessage = 'Path for the error log CSV.  Defaults to InputFile directory with timestamp.'
    )]
    [string]$ErrorLogPath,

    [Parameter(
        HelpMessage = 'Seconds to wait between site creation requests to avoid throttling.  Default: 5.'
    )]
    [ValidateRange(0, 60)]
    [int]$ThrottleDelaySeconds = 5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Helpers ---
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'SKIP')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'HH:mm:ss'
    switch ($Level) {
        'SUCCESS' { Write-Host   "[$ts][OK]   $Message" -ForegroundColor Green   }
        'SKIP'    { Write-Host   "[$ts][SKIP] $Message" -ForegroundColor DarkCyan }
        'WARN'    { Write-Warning          "[$ts][WARN] $Message"                }
        'ERROR'   { Write-Host   "[$ts][ERR]  $Message" -ForegroundColor Red     }
        default   { Write-Verbose          "[$ts][INF]  $Message"               }
    }
}

function Resolve-SiteUrl {
    <#
    .SYNOPSIS
        Converts a relative path (/sites/Name) to a full SharePoint URL.
    #>
    param ([string]$RawUrl, [string]$TenantRoot)

    if ($RawUrl -match '^https?://') {
        return $RawUrl.TrimEnd('/')
    }

    # Strip admin suffix and build the root tenant URL.
    $tenantRoot = $TenantRoot -replace '-admin\.sharepoint\.com', '.sharepoint.com'
    $tenantRoot = $tenantRoot.TrimEnd('/')

    return ($tenantRoot + '/' + $RawUrl.TrimStart('/'))
}

function Test-SiteExists {
    param ([string]$SiteUrl)
    try {
        $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop
        return $null -ne $site
    }
    catch {
        return $false
    }
}
#endregion

#region --- Default error log path ---
if (-not $ErrorLogPath) {
    $inputDir      = Split-Path -Path $InputFile -Parent
    $inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $ErrorLogPath  = Join-Path -Path $inputDir -ChildPath (
        "{0}_Errors_{1}.csv" -f $inputBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss')
    )
}
#endregion

try {
    #region --- Load and validate CSV ---
    Write-Log "Loading input file: $InputFile"
    $rows = Import-Csv -Path $InputFile -ErrorAction Stop

    if (($rows | Measure-Object).Count -eq 0) {
        Write-Log 'The input CSV contains no data rows.' -Level WARN
        exit 0
    }

    # Validate required columns exist.
    $requiredColumns = @('Title', 'Url', 'Template', 'Owner')
    $csvHeaders      = $rows[0].PSObject.Properties.Name

    foreach ($col in $requiredColumns) {
        if ($col -notin $csvHeaders) {
            throw "Required CSV column '$col' is missing.  CSV headers found: $($csvHeaders -join ', ')"
        }
    }

    # Validate Template values.
    $validTemplates = @('CommunicationSite', 'TeamSite')
    $invalidRows    = $rows | Where-Object { $_.Template -notin $validTemplates }
    if (($invalidRows | Measure-Object).Count -gt 0) {
        $badValues = ($invalidRows | Select-Object -ExpandProperty Template | Sort-Object -Unique) -join ', '
        throw "Invalid Template value(s) found: '$badValues'.  Allowed values: CommunicationSite, TeamSite"
    }

    Write-Log "Loaded $( ($rows | Measure-Object).Count ) row(s) from CSV."
    #endregion

    #region --- Connect ---
    Write-Log "Connecting to: $TenantUrl"
    Connect-PnPOnline -Url $TenantUrl -Interactive -ErrorAction Stop
    Write-Log 'Connection established.' -Level SUCCESS
    #endregion

    #region --- Process rows ---
    $total     = ($rows | Measure-Object).Count
    $current   = 0
    $created   = 0
    $skipped   = 0
    $failed    = 0
    $errorLog  = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($row in $rows) {
        $current++
        $resolvedUrl = Resolve-SiteUrl -RawUrl $row.Url -TenantRoot $TenantUrl

        Write-Progress -Activity 'Bulk Site Creation' `
                       -Status   "Site $current of $total : $($row.Title)" `
                       -PercentComplete ([int](($current / $total) * 100))

        Write-Log "--- [$current/$total] $($row.Title) ($resolvedUrl)"

        # Resume capability: skip if site already exists.
        if (Test-SiteExists -SiteUrl $resolvedUrl) {
            Write-Log "Site already exists — skipping: $resolvedUrl" -Level SKIP
            $skipped++
            continue
        }

        if ($PSCmdlet.ShouldProcess($resolvedUrl, "Create $($row.Template)")) {
            try {
                $splatParams = @{
                    Type  = $row.Template
                    Title = $row.Title
                    Url   = $resolvedUrl
                    Owner = $row.Owner
                    ErrorAction = 'Stop'
                }

                if (-not [string]::IsNullOrWhiteSpace($row.Description)) {
                    $splatParams['Description'] = $row.Description
                }

                $newSiteUrl = New-PnPSite @splatParams

                Write-Log "Created: $newSiteUrl" -Level SUCCESS
                $created++

                # Throttle between requests to reduce risk of hitting SPO rate limits.
                if ($ThrottleDelaySeconds -gt 0 -and $current -lt $total) {
                    Start-Sleep -Seconds $ThrottleDelaySeconds
                }
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Log "FAILED to create '$($row.Title)' at $resolvedUrl : $errMsg" -Level ERROR
                $failed++

                $errorLog.Add([PSCustomObject]@{
                    Title       = $row.Title
                    Url         = $resolvedUrl
                    Template    = $row.Template
                    Owner       = $row.Owner
                    Error       = $errMsg
                    Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                })
            }
        }
    }

    Write-Progress -Activity 'Bulk Site Creation' -Completed
    #endregion

    #region --- Write error log ---
    if ($errorLog.Count -gt 0) {
        $errorLog | Export-Csv -Path $ErrorLogPath -NoTypeInformation -Encoding UTF8
        Write-Log "Error log written to: $ErrorLogPath" -Level WARN
    }
    #endregion

    #region --- Summary ---
    Write-Host ''
    Write-Host '---- Bulk Site Creation Summary ----' -ForegroundColor Cyan
    Write-Host ("Total rows in CSV : {0}" -f $total)   -ForegroundColor White
    Write-Host ("Created           : {0}" -f $created)  -ForegroundColor Green
    Write-Host ("Skipped (existed) : {0}" -f $skipped)  -ForegroundColor DarkCyan
    Write-Host ("Failed            : {0}" -f $failed)   -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'White' })

    if ($failed -gt 0) {
        Write-Host "Error log         : $ErrorLogPath" -ForegroundColor Red
    }
    Write-Host ''
    #endregion
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    Write-Progress -Activity 'Bulk Site Creation' -Completed -ErrorAction SilentlyContinue
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
