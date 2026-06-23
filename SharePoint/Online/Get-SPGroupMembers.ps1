#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '1.12.0' }

<#
.SYNOPSIS
    Retrieves the members of one or more SharePoint Online groups.

.DESCRIPTION
    Get-SPGroupMembers connects to a SharePoint Online site using PnP.PowerShell
    and returns the members of the specified SharePoint group(s). Results can be
    written to the pipeline as objects, displayed in the console, or exported to
    a CSV file.

    The function supports pipeline input for both SiteUrl and GroupName, enabling
    bulk queries across multiple sites or groups in a single pass.

    AUTHENTICATION:
    This function uses Connect-PnPOnline with -Interactive by default. For
    unattended/automated scenarios, connect before calling the function using
    certificate-based auth or a Managed Identity, then pass -UseExistingConnection.

    OUTPUT OBJECT PROPERTIES:
    - SiteUrl      : The site URL queried
    - GroupName    : The SharePoint group name
    - LoginName    : The member's login name (UPN or claim string)
    - DisplayName  : The member's display name
    - Email        : The member's email address
    - IsSiteAdmin  : Whether the member is a site collection administrator
    - PrincipalType: User or SecurityGroup

.PARAMETER SiteUrl
    The full URL of the SharePoint Online site.
    Accepts pipeline input.
    Example: https://contoso.sharepoint.com/sites/MySite

.PARAMETER GroupName
    The display name of the SharePoint group to query.
    Accepts pipeline input by property name.
    Example: "MySite Members"

.PARAMETER UseExistingConnection
    When specified, the function uses the currently active PnP connection
    instead of prompting for interactive authentication. Use this when you
    have already called Connect-PnPOnline before invoking this function.

.PARAMETER ExportPath
    Optional. Full path to a CSV file where results will be exported.
    If the file already exists it will be overwritten.
    Example: C:\Reports\GroupMembers.csv

.EXAMPLE
    Get-SPGroupMembers -SiteUrl "https://contoso.sharepoint.com/sites/HR" -GroupName "HR Members"

    Description:
        Returns all members of the "HR Members" group from the HR site and
        displays them in the console.

.EXAMPLE
    Get-SPGroupMembers `
        -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
        -GroupName "HR Members" `
        -ExportPath "C:\Reports\HRMembers.csv"

    Description:
        Retrieves members and exports the result to a CSV file.

.EXAMPLE
    $groups = @(
        [PSCustomObject]@{ SiteUrl = 'https://contoso.sharepoint.com/sites/HR';      GroupName = 'HR Members' }
        [PSCustomObject]@{ SiteUrl = 'https://contoso.sharepoint.com/sites/Finance'; GroupName = 'Finance Owners' }
    )
    $groups | Get-SPGroupMembers -ExportPath "C:\Reports\AllGroups.csv"

    Description:
        Pipes multiple site/group combinations to the function and exports all
        results to a single CSV.

.EXAMPLE
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/HR" -Interactive
    Get-SPGroupMembers -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
                       -GroupName "HR Visitors" `
                       -UseExistingConnection

    Description:
        Re-uses an existing PnP connection rather than triggering a new
        interactive login prompt.

.NOTES
    AUTHOR:
        Geoff Varosky

    VERSION:
        1.0.0

    LAST UPDATED:
        2026-06-23

    REQUIREMENTS:
        - PowerShell 5.1 or 7.x
        - PnP.PowerShell 1.12.0 or later
          Install: Install-Module PnP.PowerShell -Scope CurrentUser
        - At minimum, Read access to the target SharePoint site

    REFERENCES:
        https://pnp.github.io/powershell/cmdlets/Get-PnPGroupMember.html
        https://pnp.github.io/powershell/cmdlets/Get-PnPGroup.html

    GitHub: https://github.com/VAROIndustries/SharePointYankee
#>

function Get-SPGroupMembers {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Full URL of the SharePoint Online site.'
        )]
        [ValidatePattern('^https://')]
        [string]$SiteUrl,

        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Display name of the SharePoint group to query.'
        )]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName,

        [Parameter(Mandatory = $false)]
        [switch]$UseExistingConnection,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
            $parent = Split-Path $_ -Parent
            if ($parent -and -not (Test-Path $parent)) {
                throw "Directory '$parent' does not exist. Please create it before specifying -ExportPath."
            }
            $true
        })]
        [string]$ExportPath
    )

    begin {
        # Collect all output rows so we can export at the end when -ExportPath is used.
        $allResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        $ErrorActionPreference = 'Stop'
    }

    process {
        # Connect to the site unless the caller has already established a connection.
        if (-not $UseExistingConnection) {
            try {
                Write-Verbose "Connecting to '$SiteUrl'..."
                Connect-PnPOnline -Url $SiteUrl -Interactive
            } catch {
                Write-Error "Failed to connect to '$SiteUrl': $($_.Exception.Message)"
                return
            }
        }

        # Verify the group exists before attempting to retrieve members.
        try {
            Write-Verbose "Verifying group '$GroupName' exists on '$SiteUrl'..."
            $group = Get-PnPGroup -Identity $GroupName -ErrorAction Stop
        } catch {
            Write-Warning "Group '$GroupName' not found on '$SiteUrl'. Skipping. Error: $($_.Exception.Message)"
            return
        }

        # Retrieve all members of the group.
        try {
            Write-Verbose "Retrieving members of group '$GroupName'..."
            $members = Get-PnPGroupMember -Identity $GroupName -ErrorAction Stop
        } catch {
            Write-Error "Failed to retrieve members of group '$GroupName' on '$SiteUrl': $($_.Exception.Message)"
            return
        }

        if ($null -eq $members -or $members.Count -eq 0) {
            Write-Warning "Group '$GroupName' on '$SiteUrl' has no members."
            return
        }

        Write-Verbose "Found $($members.Count) member(s) in '$GroupName'."

        foreach ($member in $members) {
            $row = [PSCustomObject]@{
                SiteUrl       = $SiteUrl
                GroupName     = $GroupName
                LoginName     = $member.LoginName
                DisplayName   = $member.Title
                Email         = $member.Email
                IsSiteAdmin   = $member.IsSiteAdmin
                PrincipalType = $member.PrincipalType.ToString()
            }

            # Emit to pipeline immediately so callers receive streaming output.
            $row

            # Also collect for potential CSV export.
            $allResults.Add($row)
        }
    }

    end {
        if ($ExportPath -and $allResults.Count -gt 0) {
            try {
                $allResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
                Write-Host "Exported $($allResults.Count) record(s) to '$ExportPath'." -ForegroundColor Green
            } catch {
                Write-Error "Failed to export results to '$ExportPath': $($_.Exception.Message)"
            }
        } elseif ($ExportPath -and $allResults.Count -eq 0) {
            Write-Warning "No results to export. CSV file was not created."
        }
    }
}
