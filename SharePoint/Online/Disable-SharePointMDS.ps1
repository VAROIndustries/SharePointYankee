<#
.SYNOPSIS
    Disables the Minimal Download Strategy (MDS) feature on one or more SharePoint sites.

.DESCRIPTION
    The Minimal Download Strategy (MDS) is a SharePoint feature (Feature ID:
    87294c72-f260-42f3-a41b-981a2ffce37a) designed to improve page load performance
    by downloading only the content that changes between page navigations.

    While MDS works well on out-of-the-box SharePoint sites, it is a well-known
    source of rendering issues in several scenarios:

        - Custom master pages or page layouts that were not built with MDS in mind
        - Third-party web parts and add-ins that manipulate the DOM directly
        - Classic SharePoint pages with certain JavaScript customizations
        - Sites migrated from SharePoint 2010 that carry legacy branding

    Symptoms of MDS conflicts include pages that render blank on first navigation but
    appear correctly after a full browser refresh (F5), broken CSS/JS after clicking
    links, or JavaScript errors referencing "asyncDeltaManager" or "SP.UI.ModalDialog".

    This script disables MDS at the site level via PnP PowerShell.  It reports the
    current state before making changes and supports -WhatIf and -Recursive options.

.PARAMETER SiteUrl
    The URL of the SharePoint site collection on which to disable MDS.
    Example: https://contoso.sharepoint.com/sites/Intranet

.PARAMETER Recursive
    When specified, also disables MDS on all subsites (webs) found beneath the
    root site of the provided URL.  Useful for site collections that have a
    deep sub-web hierarchy.

.PARAMETER Credential
    Optional PSCredential for authentication.  If omitted, PnP PowerShell will
    use the interactive browser login flow.  For app-only/certificate auth,
    call Connect-PnPOnline before invoking this script and it will reuse the
    existing connection.

.EXAMPLE
    .\Disable-SharePointMDS.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Intranet

    Disables MDS on the root web of the specified site collection.

.EXAMPLE
    .\Disable-SharePointMDS.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Intranet -Recursive

    Disables MDS on the root web and all subsites beneath it.

.EXAMPLE
    .\Disable-SharePointMDS.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Intranet -WhatIf

    Shows what would happen without making any changes.

.NOTES
    Author  : Geoff Varosky
    Module  : PnP.PowerShell 2.x+
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    MDS Feature ID: 87294c72-f260-42f3-a41b-981a2ffce37a

    The MDS feature is site-scoped (not site-collection-scoped).  Disabling it on
    the root web does not automatically affect subsites — use -Recursive for that.

    Re-enabling MDS:
        Enable-PnPFeature -Identity '87294c72-f260-42f3-a41b-981a2ffce37a' -Scope Web

    Reference:
        https://learn.microsoft.com/en-us/sharepoint/dev/general-development/minimal-download-strategy-overview
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param (
    [Parameter(
        Mandatory,
        HelpMessage = 'Full URL of the SharePoint site collection (e.g., https://contoso.sharepoint.com/sites/Intranet).'
    )]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(
        HelpMessage = 'Also process all subsites beneath the root web.'
    )]
    [switch]$Recursive,

    [Parameter(
        HelpMessage = 'Optional PSCredential.  Omit to use interactive browser login.'
    )]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# MDS feature GUID (Web-scoped)
$MdsFeatureId = '87294c72-f260-42f3-a41b-981a2ffce37a'

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
        'SUCCESS' { Write-Host   "[$ts][OK]   $Message" -ForegroundColor Green  }
        'WARN'    { Write-Warning          "[$ts][WARN] $Message"               }
        'ERROR'   { Write-Error            "[$ts][ERR]  $Message" -ErrorAction Continue }
        default   { Write-Host   "[$ts][INF] $Message" -ForegroundColor Cyan   }
    }
}

function Disable-MdsOnWeb {
    <#
    .SYNOPSIS
        Disables MDS on a single web and reports the before/after state.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string]$WebUrl
    )

    Write-Log "Processing web: $WebUrl"

    # Connect (or re-connect) to the specific web URL so Get/Set-PnP* targets it.
    try {
        if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
            Connect-PnPOnline -Url $WebUrl -Credentials $Credential -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $WebUrl -Interactive -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Failed to connect to $WebUrl : $_" -Level ERROR
        return
    }

    # Check current MDS feature state.
    $feature = Get-PnPFeature -Identity $MdsFeatureId -Scope Web -ErrorAction SilentlyContinue

    if ($null -eq $feature -or $feature.DefinitionId -eq [Guid]::Empty) {
        Write-Log "MDS is already DISABLED on: $WebUrl" -Level SUCCESS
        return
    }

    Write-Log "MDS is currently ENABLED on: $WebUrl"

    if ($PSCmdlet.ShouldProcess($WebUrl, 'Disable Minimal Download Strategy (MDS) feature')) {
        try {
            Disable-PnPFeature -Identity $MdsFeatureId -Scope Web -Force -ErrorAction Stop
            Write-Log "MDS successfully DISABLED on: $WebUrl" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to disable MDS on $WebUrl : $_" -Level ERROR
        }
    }
}
#endregion

try {
    #region --- Initial connection ---
    Write-Log "Connecting to: $SiteUrl"

    if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
        Connect-PnPOnline -Url $SiteUrl -Credentials $Credential -ErrorAction Stop
    }
    else {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
    }

    Write-Log 'Connection established.' -Level SUCCESS
    #endregion

    #region --- Collect target webs ---
    $targetUrls = [System.Collections.Generic.List[string]]::new()
    $targetUrls.Add($SiteUrl)

    if ($Recursive) {
        Write-Log 'Enumerating subsites (-Recursive specified)...'

        # Get-PnPSubWeb with -Recurse returns all descendant webs in one call.
        $subWebs = Get-PnPSubWeb -Recurse -ErrorAction Stop

        foreach ($web in $subWebs) {
            $targetUrls.Add($web.Url)
        }

        Write-Log "Found $($subWebs.Count) subsite(s).  Total webs to process: $($targetUrls.Count)"
    }
    #endregion

    #region --- Process each web ---
    $total   = $targetUrls.Count
    $current = 0

    foreach ($url in $targetUrls) {
        $current++
        Write-Progress -Activity 'Disabling MDS' `
                       -Status   "Processing $current of $total : $url" `
                       -PercentComplete ([int](($current / $total) * 100))

        Disable-MdsOnWeb -WebUrl $url
    }

    Write-Progress -Activity 'Disabling MDS' -Completed
    #endregion

    Write-Log "Done.  Processed $total web(s)." -Level SUCCESS
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
