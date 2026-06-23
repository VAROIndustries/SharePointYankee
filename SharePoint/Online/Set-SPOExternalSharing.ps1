<#
.SYNOPSIS
    Sets the external sharing level on one or more SharePoint Online sites.

.DESCRIPTION
    Configures the external sharing capability for specified SharePoint site
    collections via PnP PowerShell.  The script reports the current sharing
    level before applying any change so you have a clear before/after record
    of every site modified.

    External sharing levels (from most to least permissive):
        Anyone              — Links can be shared with anyone, no sign-in required
                              ("Anyone links" / anonymous links)
        NewAndExistingGuests — Guests must authenticate.  New external users can
                              be invited and existing guest accounts can access.
        ExistingGuests      — Only already-established guest accounts in the
                              directory can access.  No new invitations allowed.
        Disabled            — No external sharing permitted.

    Important:
        The site-level sharing setting cannot be MORE permissive than the
        tenant-level setting.  If the tenant is set to ExistingGuests, you
        cannot set an individual site to Anyone or NewAndExistingGuests.
        This script validates the requested level against the tenant policy
        and warns when a conflict is detected.

.PARAMETER SiteUrl
    One or more SharePoint site collection URLs.  Accepts pipeline input.
    Example: "https://contoso.sharepoint.com/sites/Marketing"

.PARAMETER SharingLevel
    The external sharing capability to apply.  Must be one of:
        Disabled | ExistingGuests | NewAndExistingGuests | Anyone

.PARAMETER TenantUrl
    Optional.  The SharePoint admin center URL, used only for the tenant-level
    policy validation check.  If omitted, the validation step is skipped.
    Example: https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Set-SPOExternalSharing.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Marketing -SharingLevel Disabled

    Disables external sharing on the Marketing site.

.EXAMPLE
    .\Set-SPOExternalSharing.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Marketing -SharingLevel Disabled -WhatIf

    Shows what would change without applying any settings.

.EXAMPLE
    @('https://contoso.sharepoint.com/sites/A','https://contoso.sharepoint.com/sites/B') | .\Set-SPOExternalSharing.ps1 -SharingLevel ExistingGuests

    Applies the same sharing level to multiple sites via pipeline.

.EXAMPLE
    Import-Csv sites.csv | Select-Object -ExpandProperty Url | .\Set-SPOExternalSharing.ps1 -SharingLevel Disabled

    Processes a list of URLs from a CSV file.

.NOTES
    Author  : Geoff Varosky
    Module  : PnP.PowerShell 2.x+
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+
              SharePoint Administrator role (or equivalent)
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    The PnP PowerShell SharingCapability enum values used with
    Set-PnPTenantSite map as follows:
        Disabled             → SharingCapabilities.Disabled
        ExistingExternalUserSharingOnly → ExistingGuests
        ExternalUserSharingOnly         → NewAndExistingGuests
        ExternalUserAndGuestSharing     → Anyone

    This script uses the human-friendly names defined in the -SharingLevel
    parameter and translates them to the underlying enum values internally.
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param (
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        HelpMessage = 'One or more SharePoint site collection URLs.'
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$SiteUrl,

    [Parameter(
        Mandatory,
        HelpMessage = 'External sharing level: Disabled, ExistingGuests, NewAndExistingGuests, or Anyone.'
    )]
    [ValidateSet('Disabled', 'ExistingGuests', 'NewAndExistingGuests', 'Anyone')]
    [string]$SharingLevel,

    [Parameter(
        HelpMessage = 'SharePoint admin center URL for tenant-level policy validation (optional).'
    )]
    [string]$TenantUrl
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    #region --- Sharing level translation table ---
    # Maps human-friendly names to PnP SharingCapabilities enum string values.
    $SharingLevelMap = @{
        'Disabled'               = 'Disabled'
        'ExistingGuests'         = 'ExistingExternalUserSharingOnly'
        'NewAndExistingGuests'   = 'ExternalUserSharingOnly'
        'Anyone'                 = 'ExternalUserAndGuestSharing'
    }

    # Ordered list used for permissiveness comparison (index = permissiveness level).
    $SharingPermissivenessOrder = @(
        'Disabled',
        'ExistingGuests',
        'NewAndExistingGuests',
        'Anyone'
    )
    #endregion

    #region --- Helpers ---
    function Write-Log {
        [CmdletBinding()]
        param (
            [string]$Message,
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'CHANGE')]
            [string]$Level = 'INFO'
        )
        $ts = Get-Date -Format 'HH:mm:ss'
        switch ($Level) {
            'SUCCESS' { Write-Host   "[$ts][OK]     $Message" -ForegroundColor Green   }
            'CHANGE'  { Write-Host   "[$ts][CHANGE] $Message" -ForegroundColor Cyan    }
            'WARN'    { Write-Warning          "[$ts][WARN]   $Message"                }
            'ERROR'   { Write-Host   "[$ts][ERR]    $Message" -ForegroundColor Red     }
            default   { Write-Verbose          "[$ts][INF]   $Message"                }
        }
    }

    function ConvertFrom-PnPSharingCapability {
        <#
        .SYNOPSIS
            Translates the PnP SharingCapabilities enum value back to the
            human-friendly label used by this script.
        #>
        param ([string]$PnPValue)
        switch ($PnPValue) {
            'Disabled'                          { return 'Disabled'             }
            'ExistingExternalUserSharingOnly'   { return 'ExistingGuests'       }
            'ExternalUserSharingOnly'           { return 'NewAndExistingGuests' }
            'ExternalUserAndGuestSharing'       { return 'Anyone'               }
            default                             { return $PnPValue              }
        }
    }
    #endregion

    #region --- Connect ---
    Write-Log "Connecting to SharePoint Online..."

    # Use TenantUrl if provided; otherwise connect to the first site URL.
    # The admin connection is needed for Set-PnPTenantSite regardless.
    $connectUrl = if ($TenantUrl) { $TenantUrl } else {
        # Derive admin URL from the first SiteUrl value if pipeline input is used.
        # We re-derive per item in process{} if needed; this covers non-pipeline usage.
        $SiteUrl[0] -replace '\.sharepoint\.com.*$', '-admin.sharepoint.com'
    }

    Connect-PnPOnline -Url $connectUrl -Interactive -ErrorAction Stop
    Write-Log 'Connection established.' -Level SUCCESS

    #region --- Optional tenant policy validation ---
    $tenantSharingLevel = $null

    if ($TenantUrl) {
        try {
            $tenantSettings  = Get-PnPTenant -ErrorAction Stop
            # SharingCapability is the tenant-wide external sharing setting.
            $tenantPnPValue  = $tenantSettings.SharingCapability.ToString()
            $tenantSharingLevel = ConvertFrom-PnPSharingCapability -PnPValue $tenantPnPValue

            Write-Log "Tenant external sharing level: $tenantSharingLevel"

            $requestedIndex = $SharingPermissivenessOrder.IndexOf($SharingLevel)
            $tenantIndex    = $SharingPermissivenessOrder.IndexOf($tenantSharingLevel)

            if ($requestedIndex -gt $tenantIndex) {
                Write-Log ("Requested level '{0}' is more permissive than the tenant setting '{1}'.  " +
                           "SharePoint Online will enforce the tenant limit — site setting may not apply as expected.") `
                           -f $SharingLevel, $tenantSharingLevel -Level WARN
            }
        }
        catch {
            Write-Log "Could not retrieve tenant sharing settings (non-fatal): $_" -Level WARN
        }
    }
    #endregion

    # Accumulate results for a summary at the end.
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
}

process {
    foreach ($url in $SiteUrl) {
        $url = $url.TrimEnd('/')

        Write-Log "--- Processing: $url"

        try {
            # Retrieve current site properties.
            $site = Get-PnPTenantSite -Url $url -ErrorAction Stop

            # Translate current PnP enum value to friendly name.
            $currentLevel = ConvertFrom-PnPSharingCapability -PnPValue $site.SharingCapability.ToString()

            Write-Log "  Current sharing level : $currentLevel"
            Write-Log "  Requested level       : $SharingLevel"

            if ($currentLevel -eq $SharingLevel) {
                Write-Log "  No change required — already set to '$SharingLevel'." -Level SUCCESS

                $results.Add([PSCustomObject]@{
                    SiteUrl       = $url
                    SiteTitle     = $site.Title
                    Before        = $currentLevel
                    After         = $currentLevel
                    Changed       = $false
                    Status        = 'NoChangeNeeded'
                })
                continue
            }

            if ($PSCmdlet.ShouldProcess($url, "Set external sharing from '$currentLevel' to '$SharingLevel'")) {
                $pnpValue = $SharingLevelMap[$SharingLevel]

                Set-PnPTenantSite -Url $url -SharingCapability $pnpValue -ErrorAction Stop

                Write-Log ("  Changed: '{0}' -> '{1}'" -f $currentLevel, $SharingLevel) -Level CHANGE

                $results.Add([PSCustomObject]@{
                    SiteUrl   = $url
                    SiteTitle = $site.Title
                    Before    = $currentLevel
                    After     = $SharingLevel
                    Changed   = $true
                    Status    = 'Success'
                })
            }
            else {
                # -WhatIf path
                $results.Add([PSCustomObject]@{
                    SiteUrl   = $url
                    SiteTitle = $site.Title
                    Before    = $currentLevel
                    After     = $SharingLevel
                    Changed   = $false
                    Status    = 'WhatIf'
                })
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Log "  FAILED on $url : $errMsg" -Level ERROR

            $results.Add([PSCustomObject]@{
                SiteUrl   = $url
                SiteTitle = ''
                Before    = 'Unknown'
                After     = $SharingLevel
                Changed   = $false
                Status    = "Error: $errMsg"
            })
        }
    }
}

end {
    #region --- Summary table ---
    Write-Host ''
    Write-Host '---- External Sharing Change Summary ----' -ForegroundColor Cyan

    $results | Format-Table -AutoSize -Property SiteTitle, SiteUrl, Before, After, Status

    $changed = ($results | Where-Object { $_.Changed }).Count
    $errors  = ($results | Where-Object { $_.Status -like 'Error:*' }).Count

    Write-Host ("Sites processed : {0}" -f $results.Count)        -ForegroundColor White
    Write-Host ("Changed         : {0}" -f $changed)               -ForegroundColor Green
    Write-Host ("No change needed: {0}" -f ($results | Where-Object { $_.Status -eq 'NoChangeNeeded' }).Count) -ForegroundColor White
    Write-Host ("Errors          : {0}" -f $errors)                -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'White' })
    Write-Host ''
    #endregion

    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
